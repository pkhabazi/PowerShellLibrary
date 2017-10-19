$wshell = New-Object -ComObject Wscript.Shell
#Login AzureRM
try {
    Try {
        Get-AzureRmSubscription -ErrorAction Continue
    } Catch {
        if ($_ -like "*Login-AzureRmAccount to login*") {
            Write-Host 'Start loging in Azure' -ForegroundColor Yellow
          Login-AzureRmAccount
          Write-Host 'Successfully logged id' -ForegroundColor Green
        }
    }
    Write-Host 'Select Subscription' -ForegroundColor Yellow
    Set-AzureRmContext -SubscriptionName $(if ((Get-AzureRmSubscription).count -gt '1') { ((Get-AzureRmSubscription).Name | Out-GridView -Title "Select Azure Subscription" -PassThru) } else {(Get-AzureRmSubscription).Name})
    Write-Host 'Subscription successfully selected' -ForegroundColor Green
} catch {
    Write-Host $_.Exception.Message
    break
}
# Select Azure Resource Group and VM
try {
    Write-Host 'Select Resource Group and VM'
    $RgName = $((Get-AzureRmResourceGroup).ResourceGroupName | Out-GridView -Title "Select the Resource Group which contains the vm" -PassThru)
    $VMName = if ((Get-AzureRmVM -ResourceGroupName $RgName -ErrorAction Continue).count -eq '0') {$wshell.Popup("No VM found in the current Resource Group",0,"Error",0x0); break} else {(Get-AzureRmVM -ResourceGroupName $RgName).Name | Out-GridView -Title "Select the VM you want to move" -PassThru}
    $RgNameNew = $((Get-AzureRmResourceGroup).ResourceGroupName | Out-GridView -Title "Select the Resource Group where you want to move the VM to" -PassThru)
    Write-Host 'Resources successfully selected' -ForegroundColor Green
} catch {
    Write-Host  $_.Exception.Message        
}

try {
    Write-Host 'Start collecting VM Information' -ForegroundColor Yellow
    ## Collect Information for redeploying the VM
    $vmConfigExisting = Get-AzureRmvm -VMName $vmName -ResourceGroupName $RgName
    $NetworkInterface = Get-AzureRmNetworkInterface | Where-Object {$_.VirtualMachine -ne $null} | Where-Object {$_.VirtualMachine.Id -eq $vmConfigExisting.Id}
    $vmStatus = (Get-AzureRmvm -VMName $vmName -ResourceGroupName $RgName -Status).statuses[1].DisplayStatus
    ## VM Info
    $location = $vmConfigExisting.Location
    $NickName = ($NetworkInterface.Id).Split('/')| Select-Object -Last 1
    $NickId =  $vmConfigExisting.NetworkProfile.NetworkInterfaces.id
    $osDiskName = $vmConfigExisting.StorageProfile.OsDisk.Name
    $VMSize = $vmConfigExisting.HardwareProfile.VmSize
    $IPaddress = $NetworkInterface.IpConfigurations.PrivateIpAddress
    $subnet = $NetworkInterface.IpConfigurations.Subnet.Id
    ## Snapshot and Storage variables
    $StorageAccountName = "tempmigstor$(Get-Date -Format sMmmshh)"    
    $storageAccountType = "StandardLRS"
    $snapshotName = $VMName + "-Snapshot"              
    $destContainer = "vhds"
    $blobName = $snapshotName + ".vhd"
    $osDiskUri = "https://$StorageAccountName.blob.core.windows.net/$destContainer/$blobName"
    Write-Host "Successfully collectted all the information" -ForegroundColor Green
} catch {
    Write-Host  $_.Exception.Message
    break
}
function Create-Snapshot {
    try {
        ## Shutdown VM before making Snapshot
        if ($vmStatus -like "VM running") {
            Write-Host 'VM is running, stopping VM..' -ForegroundColor Yellow
            Stop-AzureRmVM -Name $VMName -ResourceGroupName $RgName -Force
            Write-Host 'VM successfully stopped' -ForegroundColor Green
        }
        ### Create Snapshot
        try {
            Write-Host 'Start creating snapshot' -ForegroundColor Yellow
            $Disk = Get-AzureRmDisk -ResourceGroupName $RgName -DiskName $($vmConfigExisting.StorageProfile.OsDisk.Name)
            $storageAccountType = $Disk.AccountType
            $Snapshot =  New-AzureRmSnapshotConfig -SourceUri $Disk.Id -CreateOption Copy -Location $location
            New-AzureRmSnapshot -Snapshot $Snapshot -SnapshotName $snapshotName -ResourceGroupName $RgNameNew
            $absoluteUri = (Grant-AzureRmSnapshotAccess -ResourceGroupName $RgNameNew -SnapshotName $snapshotName -Access 'Read' -DurationInSecond 3600).AccessSAS
            Write-Host 'Snapshot successfully created' -ForegroundColor Green
        } catch {
            Write-Host $_.Exception.Message
            break
        }
        ### Create Storage Account if not exist
        try {
            if (Get-AzureRmStorageAccount -ResourceGroupName $RgNameNew -Name $StorageAccountName -ErrorAction SilentlyContinue){
                Write-Host "Storage Account Exist" -ForegroundColor Green
               $ctx = (Get-AzureRmStorageAccount -ResourceGroupName $RgNameNew -Name $StorageAccountName).Context     
           } else {
               Write-Host "Storage Account doesnt't exist and will be created!" -ForegroundColor Yellow
               $NewStoragAccount = New-AzureRmStorageAccount -ResourceGroupName $RgNameNew -Name $StorageAccountName -Location $location -SkuName "Standard_LRS" -ErrorAction Stop
               $ctx = $NewStoragAccount.Context
               Write-Host 'Storage Account successfully created' -ForegroundColor Green
           }  
        } catch {
            Write-Host $_.Exception.Message
            break
        }
        ### Copy Snapshot to Azure Blob
        Write-Host 'start copying snapshot to Azure Blob' -ForegroundColor Yellow
        if (! (Get-AzureStorageContainer -Name $destContainer -Context $ctx -ErrorAction SilentlyContinue)){
            New-AzureStorageContainer -Name $destContainer -Context $ctx -Permission blob -ErrorAction SilentlyContinue
        }
        $destContext = New-AzureStorageContext -StorageAccountName $storageAccountName -StorageAccountKey ((Get-AzureRmStorageAccountKey -ResourceGroupName $RgNameNew -Name $StorageAccountName)[0].Value)
        Start-AzureStorageBlobCopy -AbsoluteUri $absoluteUri -DestContainer $destContainer -DestContext $destContext -DestBlob $blobName
        $CopyStatus = Get-AzureStorageBlobCopyState -Blob $blobName -Container $destContainer -Context $ctx
        while ($CopyStatus.status -like 'Pending') {
            try {
                $CopyStatus = Get-AzureStorageBlobCopyState -Blob $blobName -Container $destContainer -Context $ctx
                Start-Sleep -Seconds 30
            } catch {
                Write-Host $_.Exception.Message
            }
        }
        Write-Host 'Succesvoll copied content to Azure Blob' -ForegroundColor Green
    } catch {
        Write-Host $_.Exception.Message
    }
}
function Create-NewAzureRMVM {
    try {
        Write-Host 'Start removing VM' -ForegroundColor Yellow
        Remove-AzureRmVM -Name $VMName -Force -ResourceGroupName $RgName
        Remove-AzureRmNetworkInterface -Name $NickName -ResourceGroupName $RgName -Force
        Remove-AzureRmDisk -ResourceGroupName $RgName -DiskName $osDiskName -Force
        Write-Host 'VM successfully removed' -ForegroundColor Green
    } catch {
        Write-Host $_.Exception.Message
    }
    Try {
        Write-Host 'Start creating VM' -ForegroundColor Yellow
        $IPconfig = New-AzureRmNetworkInterfaceIpConfig -Name 'IPConfig1' -PrivateIpAddressVersion 'IPv4' -PrivateIpAddress $IPaddress -SubnetId $subnet
        $nic = New-AzureRmNetworkInterface -Name $NickName -ResourceGroupName $RgNameNew -Location $location -IpConfiguration $IPconfig -Confirm:$false -Force
        $vmConfig = New-AzureRmVMConfig -VMName $vmName -VMSize $VMSize
        $vm = Add-AzureRmVMNetworkInterface -VM $vmConfig -Id $nic.Id
        $osDisk = New-AzureRmDisk -DiskName $osDiskName -Disk (New-AzureRmDiskConfig -AccountType $storageAccountType -Location $location -CreateOption Import -SourceUri $osDiskUri) -ResourceGroupName $RgNameNew
        $vm = Set-AzureRmVMOSDisk -VM $vm -ManagedDiskId $osDisk.Id -StorageAccountType $storageAccountType -DiskSizeInGB $($vmConfigExisting.StorageProfile.OsDisk.DiskSizeGB) -CreateOption Attach -Windows
        $vm = Set-AzureRmVMBootDiagnostics -VM $vm -disable
        #Create the new VM
        New-AzureRmVM -ResourceGroupName $RgNameNew -Location $location -VM $vm
        Write-Host 'VM successfully created' -ForegroundColor Green
    } catch {
        Write-Host $_.Exception.Message
    }
}
function Clean-Migration {
    $clean = $($wshell.Popup("Operation Completed, do you want to clean the snapshot",0,"Done",0x4))
    if ($clean -eq '6') {
        Write-Host 'Start Cleaning' -ForegroundColor Yellow
        try {
            Remove-AzureRmStorageAccount -ResourceGroupName $RgNameNew -Name $StorageAccountName -Force
            Write-Host 'Azure Storage Accpunt successfully removed'  -ForegroundColor Green 
            Revoke-AzureRmSnapshotAccess -ResourceGroupNam $RgNameNew -SnapshotName $snapshotName
            Write-Host 'Snapshot Acces key successfully revoked' -ForegroundColor Green            
            Remove-AzureRmSnapshot -ResourceGroupName $RgNameNew -SnapshotName $snapshotName -Force
            Write-Host 'Azure Snapshhot successfully removed' -ForegroundColor Green
        } catch {
            Write-Host $_.Exception.Message
        }
    } else {
        Write-Host 'No Cleaning' -ForegroundColor Gray
    }
}

Create-Snapshot
Create-NewAzureRMVM
Clean-Migration
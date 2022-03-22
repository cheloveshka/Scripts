# Получаем из контейнера нашего региона все имена ПК с нужными нам ОС в переменную
$PC_List = Get-ADComputer -SearchBase "OU=Moscow, OU=Workstations, DC=ad, DC=kdl-test, DC=ru" -Filter {OperatingSystem -notlike '*Windows 10*' -and  OperatingSystem -notlike '*Windows Server*'} | Select -ExpandProperty Name

# Импортируем список с DNS именами ПК
#$PC_List = Get-Content -Path C:\Project\PowerShell\Exchange\Export\DNS.txt 

# Создаём пустой массив
$results = @()

# Обрабатываем построчно, для каждого PC
ForEach ($PC in $PC_List)
 { $obj=new-object psobject
    
    $PC_Type = ""
	$PC_Processor = ""
	$PC_Motherboard = ""
    
	# Выполняем получение данных о производителе компьютера
	$PC_Type = get-wmiobject Win32_ComputerSystem -computer $PC.HostName 
	# Записываю данные в строку по значениям 
    #$obj | Add-Member -MemberType NoteProperty -Name PC_Name -Value ($PC_Type.Name)
	$obj | Add-Member -MemberType NoteProperty -Name PC_Manufacturer -Value ($PC_Type.Manufacturer)
	$obj | Add-Member -MemberType NoteProperty -Name PC_Model -Value ($PC_Type.Model)
	$obj | Add-Member -MemberType NoteProperty -Name RAM -Value ($PC_Type.TotalPhysicalMemory)

	# Выполняем получение данных о процессоре компьютера
    $PC_Processor = get-wmiobject Win32_processor -computer $PC
	# Записываю данные в строку по значениям 
	$obj | Add-Member -MemberType NoteProperty -Name Processor -Value ($PC_Processor.Name)
	$obj | Add-Member -MemberType NoteProperty -Name Cores -Value ($PC_Processor.numberOfCores)
	$obj | Add-Member -MemberType NoteProperty -Name LogicalProcessors -Value ($PC_Processor.NumberOfLogicalProcessors)
	
	# Выполняем получение данных о материнской плате компьютера
    $PC_Motherboard = get-wmiobject Win32_baseboard -computer $PC
	# Записываю данные в строку по значениям 
	$obj | Add-Member -MemberType NoteProperty -Name MotherBoard -Value ($PC_Motherboard.Product)
	$obj | Add-Member -MemberType NoteProperty -Name Manufacturer -Value ($PC_Motherboard.Manufacturer)
	#$obj | Add-Member -MemberType NoteProperty -Name SN -Value ($PC_Motherboard.SerialNumber)
	#$obj | Add-Member -MemberType NoteProperty -Name Version -Value ($PC_Motherboard.Version)
	
    # Формируем массив объектов
    $results +=$obj
}
# Выводим результат в файл
$results | Select-Object -Property PC_Name,PC_Manufacturer,PC_Model,RAM,Processor,Cores,LogicalProcessors,MotherBoard,Manufacturer,SN,Version | Export-Csv C:\Project\PowerShell\Result\PC_Info.csv -Encoding UTF8 -Delimiter "," -NoTypeInformation -Force
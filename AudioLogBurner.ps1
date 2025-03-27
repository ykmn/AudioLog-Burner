# Объявление параметра скрипта
param (
    [string]$YearMonth
)

# Установка кодировки на UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Максимальный размер:
$maxSize = 4.5GB
# DVD-R: 4.5GB
# DVD-R Dual Layer: 8.1GB
# Blu-Ray: 24.5GB
# Blu-Ray Double Layer: 49.5GB


# Путь к текущей папке скрипта
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$imgBurnPath = Join-Path -Path $scriptDir -ChildPath "ImgBurn.exe"
$configFile = "config.json"

Write-Host ""
Write-Host "AudioLogBurner.ps1                            v1.02 2025-03-26 Roman Ermakov <r.ermakov@emg.fm>"
Write-Host "Скрипт копирует папки с файлами контроля эфира в соответствии с файлом конфигурации $configFile"
Write-Host "создает ISO-образы и записывает их на CD/DVD/BR для передачи в Гостелерадиофонд. `n"
Write-Host "По умолчанию скрипт ищет папки за предыдущий месяц. Если нужно скопировать конкретный месяц,"
Write-Host "Запустите скрипт с параметром: " -NoNewline
Write-Host ".\AudioLogBurner.ps1 -YearMonth 2024-11" -ForegroundColor Yellow
Write-Host ""

# Проверка наличия ImgBurn.exe в папке со скриптом
if (-Not (Test-Path $imgBurnPath)) {
    Write-Host "ImgBurn.exe не найден в папке скрипта. Начинаю загрузку установщика с https://imgburn.com" -ForegroundColor Yellow
    $imgBurnUrl = "https://download.imgburn.com/SetupImgBurn_2.5.8.0.exe"
    $installerPath = Join-Path -Path $scriptDir -ChildPath "SetupImgBurn_2.5.8.0.exe"
    try {
        Write-Host "Скачиваю $imgBurnUrl..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $imgBurnUrl -OutFile $installerPath
        Write-Host "Установщик ImgBurn успешно скачан: $installerPath" -ForegroundColor Green
        Write-Host "Пожалуйста, запустите $installerPath вручную, установите ImgBurn, а затем переместите ImgBurn.exe в папку со скриптом ($scriptDir). После этого перезапустите скрипт." -ForegroundColor Yellow
        exit 1
    }
    catch {
        Write-Host "Ошибка при скачивании установщика ImgBurn: $_" -ForegroundColor Red
        exit 1
    }
}
else {
    Write-Host "ImgBurn.exe найден в $scriptDir. Продолжаю выполнение скрипта." -ForegroundColor Green
}

# Чтение конфигурационного файла
if (-Not (Test-Path $configFile)) {
    Write-Host "Конфигурационный json-файл '$configFile' не найден." -ForegroundColor Red
    Write-Host "Создайте его в виде массива:"
    Write-Host '{'
    Write-Host '    "paths": ['
    Write-Host '        {'
    Write-Host '            "station": "Europa Plus",'
    Write-Host '            "source": "\\\\SERVER\\LOGGER\\01 Europa FM",'
    Write-Host '            "destination": "D:\\LOGGER\\Europa Plus",'
    Write-Host '            "drive": "K:"'
    Write-Host '        },'
    Write-Host '        {'
    Write-Host '            "station": "Retro FM",'
    Write-Host '            "source": "\\\\SERVER\\LOGGER\\02 Retro FM",'
    Write-Host '            "destination": "D:\\LOGGER\\Retro FM",'
    Write-Host '            "drive": "L:"'
    Write-Host '        }'
    Write-Host '    ]'
    Write-Host '}'
    Write-Host 'Не забывайте, что слэши должны быть двойными.'
    Write-Host 'drive = буква записывающего привода, могут быть разными.'
    exit 1
}

# Загрузка конфигурации из JSON
try {
    $config = Get-Content $configFile | ConvertFrom-Json
}
catch {
    Write-Host "Ошибка загрузки конфигурации из $configFile : $_" -ForegroundColor Red
    exit 1
}

# --- Часть 1: Копирование файлов за указанный или прошлый месяц ---

$currentDate = Get-Date

if ($YearMonth) {
    if ($YearMonth -match '^\d{4}-\d{2}$') {
        $year, $month = $YearMonth.Split('-')
        if ([int]$month -ge 1 -and [int]$month -le 12) {
            $firstDay = Get-Date -Year $year -Month $month -Day 1
            $lastDay = $firstDay.AddMonths(1).AddDays(-1)
            Write-Host "Копирование файлов за указанный месяц: $YearMonth" -ForegroundColor Cyan
        } else {
            Write-Host "Ошибка: месяц должен быть в диапазоне от 01 до 12 (например, 2025-03)" -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "Ошибка: параметр должен быть в формате yyyy-mm (например, 2025-03)" -ForegroundColor Red
        exit 1
    }
} else {
    $firstDay = (Get-Date -Year $currentDate.Year -Month $currentDate.Month -Day 1).AddMonths(-1)
    $lastDay = $firstDay.AddDays((Get-Date -Year $firstDay.Year -Month $firstDay.Month -Day 1).AddMonths(1).AddDays(-1).Day - 1)
    Write-Host "Копирование файлов за предыдущий месяц: $($firstDay.ToString('yyyy-MM'))" -ForegroundColor Cyan
}

# Форматируем даты для поиска папок
$folderNames = @()
$currentDate = $firstDay
while ($currentDate -le $lastDay) {
    $folderNames += $currentDate.ToString("yyyy-MM-dd")
    $currentDate = $currentDate.AddDays(1)
}

# Обработка каждой пары source и destination для копирования
foreach ($path in $config.paths) {
    $currentSource = $path.source
    $currentDestination = $path.destination
    $currentStation = $path.station
    Write-Host "`nПолучение файлов: $currentStation " -ForegroundColor Green
    if (-Not (Test-Path $currentDestination)) {
        Write-Host "Папка назначения '$currentDestination' не найдена. Создаем папку." -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $currentDestination -Force | Out-Null
    }

    $folders = Get-ChildItem -Path $currentSource -Directory | Where-Object {
        $folderNames -contains $_.Name
    }

        if ($folders.Count -gt 0) {
        foreach ($folder in $folders) {
            $sourceFolderPath = $folder.FullName
            $filesToCopy = Get-ChildItem "$sourceFolderPath\*" -Recurse
            $totalFiles = $filesToCopy.Count
            $fileIndex = 0

            #Write-Host "Копирование из '$sourceFolderPath' в '$currentDestination'" -ForegroundColor Green
            Write-Host "$($folder.Name) " -NoNewline

            $destinationFolderPath = Join-Path -Path $currentDestination -ChildPath $folder.Name
            # Write-Host "Создаю директорию $destinationFolderPath" -BackgroundColor Green -ForegroundColor Black
            New-Item -ItemType Directory -Path $destinationFolderPath -Force -ErrorAction SilentlyContinue | Out-Null
            
            foreach ($file in $filesToCopy) {
                $destinationFilePath = Join-Path -Path $destinationFolderPath -ChildPath $file.Name

                if (Test-Path $destinationFilePath) {
                    $destinationFileInfo = Get-Item $destinationFilePath
                    if ($file.Length.CompareTo($destinationFileInfo.Length) -le 0) {
                        continue
                    }
                }

                $destinationDir = Split-Path -Path $destinationFilePath -Parent
                if (-Not (Test-Path $destinationDir)) {
                    New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
                    Write-Host "`n Создана директория назначения '$destinationDir'" -ForegroundColor Yellow
                }

                Copy-Item -Path $file.FullName -Destination $destinationFilePath -Force
                $fileIndex++
                Write-Progress -Activity "Копирование файлов" -Status "$fileIndex из $totalFiles" -PercentComplete (($fileIndex / $totalFiles) * 100)
            }

            # Write-Host "Копирование из '$sourceFolderPath' завершено.`n" -ForegroundColor Green
        }
    } else {
        Write-Host "Нет папок в '$currentSource' за указанный месяц." -ForegroundColor Red
    }
}
Write-Host ""


# --- Часть 2: Обработка файлов и создание ISO ---

$pathIndex = 0
foreach ($path in $config.paths) {
    $pathIndex++
    $currentDestination = $path.destination
    $drive = $path.drive
    $station = $path.station
    
    Write-Host ""
    Write-Host "`nОбработка: $station" -BackgroundColor DarkGreen -ForegroundColor Black
    Write-Host ""

    if (-Not (Test-Path $currentDestination)) {
        Write-Host "Папка назначения '$currentDestination' не найдена." -ForegroundColor Yellow
        continue
    }

    $allFolders = Get-ChildItem -Path $currentDestination -Directory
    Write-Host "`n"
    Write-Host "Все папки в '$currentDestination' перед фильтрацией:" -ForegroundColor Cyan
    $allFolders | ForEach-Object { Write-Host "$($_.Name) " -NoNewline }

    $dateFolders = $allFolders | Where-Object {
        $folderNames -contains $_.Name -and $_.Name -match '^\d{4}-\d{2}-\d{2}$'
    }
    
    Write-Host "`n"
    Write-Host "Отфильтрованные папки в '$currentDestination' для создания ISO:" -ForegroundColor Yellow
    if ($dateFolders.Count -gt 0) {
        $dateFolders | ForEach-Object { Write-Host "$($_.Name) " -NoNewline }
    } else {
        Write-Host " - Нет подходящих папок" -ForegroundColor Red
    }

    Write-Host "`n"
    $files = $dateFolders | Get-ChildItem -File -Recurse
    $totalSize = 0
    $currentBatch = @()
    $batchIndex = 1
    $fileCount = $files.Count
    $processedFiles = 0

    if ($dateFolders.Count -eq 0) {
        Write-Host "Нет папок в '$currentDestination' за указанный месяц для создания ISO." -ForegroundColor Yellow
        continue
    }

    Write-Progress -Activity "Обработка папки '$currentDestination'" `
                  -Status "$pathIndex из $($config.paths.Count) папок" `
                  -PercentComplete (($pathIndex / $config.paths.Count) * 100) `
                  -CurrentOperation "Инициализация"

    $filesByFolder = $files | Group-Object { Split-Path -Path $_.DirectoryName -Leaf }

    foreach ($folderGroup in $filesByFolder) {
        $folderFiles = $folderGroup.Group
        $folderSize = ($folderFiles | Measure-Object -Property Length -Sum).Sum

        Write-Progress -Activity "Формирование партий в '$currentDestination'" `
                      -Status "Файл $processedFiles из $fileCount (Партия $batchIndex)" `
                      -PercentComplete (($processedFiles / $fileCount) * 100) `
                      -ParentId 1 `
                      -CurrentOperation "Анализ размера папки $($folderGroup.Name)"

        if ($totalSize + $folderSize -gt $maxSize -and $currentBatch.Count -gt 0) {
            $subFolderPath = Join-Path -Path $currentDestination -ChildPath "Batch_$batchIndex"
            $firstFileDate = Split-Path -Path $currentBatch[0].DirectoryName -Leaf
            $yearMonth = $firstFileDate.Substring(0, 7)
            $destinationName = Split-Path -Path $currentDestination -Leaf
            $isoFileName = Join-Path -Path $currentDestination -ChildPath "${destinationName}_${yearMonth}_Batch$batchIndex.iso"

            if (-Not (Test-Path $isoFileName)) {
                New-Item -ItemType Directory -Path $subFolderPath -Force | Out-Null
                $moveIndex = 0
                $batchSize = $currentBatch.Count
                $folderStructure = @{}

                foreach ($batchFile in $currentBatch) {
                    $moveIndex++
                    $parentFolder = Split-Path -Path $batchFile.DirectoryName -Leaf
                    $destinationFolder = Join-Path -Path $subFolderPath -ChildPath $parentFolder

                    if (-Not (Test-Path $destinationFolder)) {
                        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
                    }

                    Write-Progress -Activity "Перенос файлов в Batch_$batchIndex" `
                                  -Status "Файл $moveIndex из $batchSize" `
                                  -PercentComplete (($moveIndex / $batchSize) * 100) `
                                  -ParentId 2 `
                                  -CurrentOperation "Перенос файла: $($batchFile.Name)"

                    Move-Item -Path $batchFile.FullName -Destination $destinationFolder -Force

                    if (-not $folderStructure.ContainsKey($parentFolder)) {
                        $folderStructure[$parentFolder] = @()
                    }
                    $folderStructure[$parentFolder] += $batchFile.Name
                }

                Write-Progress -Activity "Создание ISO для Batch_$batchIndex" `
                              -Status "Запуск ImgBurn" `
                              -PercentComplete 0 `
                              -ParentId 1 `
                              -CurrentOperation "Формирование образа: $isoFileName"

                Start-Process "$imgBurnPath" -ArgumentList "/YES /MODE BUILD /BUILDMODE IMAGEFILE /SRC `"$subFolderPath`" /DEST `"$isoFileName`" /START /CLOSE /OVERWRITE YES /ROOTFOLDER YES /NOIMAGEDETAILS /VOLUMELABEL `"Batch_$batchIndex`"" -Wait

                # Сохранение содержимого Batch в текстовый файл
                $txtFileName = Join-Path -Path $currentDestination -ChildPath "${destinationName}_${yearMonth}_Batch${batchIndex}.txt"
                $content = "Содержимое ${destinationName}_${yearMonth}_Batch${batchIndex}:`n"
                foreach ($folder in $folderStructure.Keys | Sort-Object) {
                    $content += "`nПапка: $folder`n"
                    foreach ($fileName in $folderStructure[$folder]) {
                        $content += "  - $fileName`n"
                    }
                }
                Set-Content -Path $txtFileName -Value $content -Encoding UTF8

                # Возврат папок из Batch
                $subFolders = Get-ChildItem -Path $subFolderPath -Directory
                $subFolderCount = $subFolders.Count
                $subFolderIndex = 0
                #Write-Host "`nВозврат папок из Batch_$($batchIndex):" -ForegroundColor Yellow
                foreach ($subFolder in $subFolders) {
                    $subFolderIndex++
                    Write-Progress -Activity "Возврат папок из Batch_$batchIndex в '$currentDestination'" `
                                  -Status "Папка $subFolderIndex из $subFolderCount" `
                                  -PercentComplete (($subFolderIndex / $subFolderCount) * 100) `
                                  -ParentId 1 `
                                  -CurrentOperation "Возврат папки: $($subFolder.Name)"

                    $destinationFolder = Join-Path -Path $currentDestination -ChildPath $subFolder.Name
                    if (-Not (Test-Path $destinationFolder)) {
                        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
                    }
                    Get-ChildItem -Path $subFolder.FullName -File | Move-Item -Destination $destinationFolder -Force
                    #Write-Host " $($subFolder.Name)"  -NoNewline
                }
                Remove-Item -Path $subFolderPath -Force -Recurse
            }
            else {
                Write-Host "ISO-файл $isoFileName уже существует, пропускаем создание и перенос файлов." -ForegroundColor Yellow
            }

            Write-Progress -Activity "Запись ISO на диск для Batch_$batchIndex" `
                          -Status "Запуск ImgBurn для записи (с тестовым режимом)" `
                          -PercentComplete 0 `
                          -ParentId 1 `
                          -CurrentOperation "Запись на привод: $drive"
            Start-Process "$imgBurnPath" -ArgumentList "/MODE WRITE /SRC `"$isoFileName`" /DEST `"$drive`" /START /CLOSE /EJECT /VERIFY YES /WAITFORMEDIA" -Wait

            $currentBatch = @()
            $totalSize = 0
            $batchIndex++
        }

        foreach ($file in $folderFiles) {
            $currentBatch += $file
            $totalSize += $file.Length
            $processedFiles++
        }
    }

    if ($currentBatch.Count -gt 0) {
        $subFolderPath = Join-Path -Path $currentDestination -ChildPath "Batch_$batchIndex"
        $firstFileDate = Split-Path -Path $currentBatch[0].DirectoryName -Leaf
        $yearMonth = $firstFileDate.Substring(0, 7)
        $destinationName = Split-Path -Path $currentDestination -Leaf
        $isoFileName = Join-Path -Path $currentDestination -ChildPath "${destinationName}_${yearMonth}_Batch$batchIndex.iso"

        if (-Not (Test-Path $isoFileName)) {
            New-Item -ItemType Directory -Path $subFolderPath -Force | Out-Null
            $moveIndex = 0
            $batchSize = $currentBatch.Count
            $folderStructure = @{}

            foreach ($batchFile in $currentBatch) {
                $moveIndex++
                $parentFolder = Split-Path -Path $batchFile.DirectoryName -Leaf
                $destinationFolder = Join-Path -Path $subFolderPath -ChildPath $parentFolder

                if (-Not (Test-Path $destinationFolder)) {
                    New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
                }

                Write-Progress -Activity "Перенос файлов в Batch_$batchIndex" `
                              -Status "Файл $moveIndex из $batchSize" `
                              -PercentComplete (($moveIndex / $batchSize) * 100) `
                              -ParentId 2 `
                              -CurrentOperation "Перенос файла: $($batchFile.Name)"

                Move-Item -Path $batchFile.FullName -Destination $destinationFolder -Force

                if (-not $folderStructure.ContainsKey($parentFolder)) {
                    $folderStructure[$parentFolder] = @()
                }
                $folderStructure[$parentFolder] += $batchFile.Name
            }

            Write-Progress -Activity "Создание ISO для Batch_$batchIndex" `
                          -Status "Запуск ImgBurn" `
                          -PercentComplete 0 `
                          -ParentId 1 `
                          -CurrentOperation "Формирование образа: $isoFileName"

            Start-Process "$imgBurnPath" -ArgumentList "/YES /MODE BUILD /BUILDMODE IMAGEFILE /SRC `"$subFolderPath`" /DEST `"$isoFileName`" /START /CLOSE /OVERWRITE YES /ROOTFOLDER YES /NOIMAGEDETAILS /VOLUMELABEL `"Batch_$batchIndex`"" -Wait

            $txtFileName = Join-Path -Path $currentDestination -ChildPath "${destinationName}_${yearMonth}_Batch${batchIndex}.txt"
            $content = "Содержимое ${destinationName}_${yearMonth}_Batch${batchIndex}:`n"
            foreach ($folder in $folderStructure.Keys | Sort-Object) {
                $content += "`nПапка: $folder`n"
                foreach ($fileName in $folderStructure[$folder]) {
                    $content += "  - $fileName`n"
                }
            }
            Set-Content -Path $txtFileName -Value $content -Encoding UTF8

            $subFolders = Get-ChildItem -Path $subFolderPath -Directory
            $subFolderCount = $subFolders.Count
            $subFolderIndex = 0

            foreach ($subFolder in $subFolders) {
                $subFolderIndex++
                Write-Progress -Activity "Возврат папок из Batch_$batchIndex в '$currentDestination'" `
                              -Status "Папка $subFolderIndex из $subFolderCount" `
                              -PercentComplete (($subFolderIndex / $subFolderCount) * 100) `
                              -ParentId 1 `
                              -CurrentOperation "Возврат папки: $($subFolder.Name)"

                $destinationFolder = Join-Path -Path $currentDestination -ChildPath $subFolder.Name
                if (-Not (Test-Path $destinationFolder)) {
                    New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
                }
                Get-ChildItem -Path $subFolder.FullName -File | Move-Item -Destination $destinationFolder -Force
                #Write-Host "Перемещено содержимое папки '$($subFolder.Name)' в '$destinationFolder'" -ForegroundColor Cyan
                Write-Host " $($subFolder.Name)" -NoNewline
            }

            Remove-Item -Path $subFolderPath -Force -Recurse
        }
        else {
            Write-Host "`nISO-файл '$isoFileName' уже существует, пропускаем создание и перенос файлов." -ForegroundColor Yellow
        }

        Write-Progress -Activity "Запись ISO на диск для Batch_$batchIndex" `
                      -Status "Запуск ImgBurn с тестовым режимом для записи 2 экземпляров дисков" `
                      -PercentComplete 0 `
                      -ParentId 1 `
                      -CurrentOperation "Запись на привод: $drive"
        Start-Process "$imgBurnPath" -ArgumentList "/COPIES 2 /MODE WRITE /SRC `"$isoFileName`" /DEST `"$drive`" /START /CLOSE /EJECT /VERIFY YES /WAITFORMEDIA /TESTMODE YES" -Wait
    }

    Write-Host "`nВсе файлы из '$currentDestination' за указанный месяц были обработаны." -ForegroundColor Green
}
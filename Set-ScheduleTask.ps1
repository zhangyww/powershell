$taskName = "newTask1"
$taskDescription =  "newTask1 Description"
$taskCommand = "powershell"
$taskScript = "test.ps1"
# $taskArgs = "-WindowStyle Hidden -NoInteractive -ExecutionPolicy unrestricted -file $taskScript"
$taskArgs = @"
-WindowStyle Hidden -ExecutionPolicy unrestricted -Command "echo 'hello'>> E:\a.txt"
"@
$taskStartTime = [datetime]::Now.AddMinutes(1)

# 创建Schedule.Service的COM对象
$service = new-object -ComObject("Schedule.Service")
$service.connect()

# 获取定义定时任务的目录
$rootFolder = $service.GetFolder("\")

# 创建一个定时任务，参数0是保留参数，必须是0
$taskDefinition = $service.NewTask(0)
# 设置定时任务的作者、描述和创建日期
$taskDefinition.RegistrationInfo.Description = "$taskDescrition"
$taskDefinition.RegistrationInfo.Author = "Administrator"

# 设置定时任务(Enable是否可用)、可用时开启等
$taskDefinition.Settings.Enabled = $true
$taskDefinition.Settings.StartWhenAvailable = $true
$taskDefinition.Settings.Hidden = $false

# 设置触发器
# Create中 1表示计划时 8表示当开机时
$trigger = $taskDefinition.Triggers.Create(1)
$trigger.Enabled = $true
$trigger.StartBoundary = $taskStartTime.ToString("yyyy-MM-dd'T'HH:mm:ss")

# duration格式
# PnYnMnDTnHnMnS 最短为1分钟，不设置为无穷
# $trigger.Repetition.Duration = 

# interval格式
# PnDTnHnMnS 最长31天，最短1分钟
$trigger.Repetition.Interval = "PT1M"

# 创建和设置定时任务的操作
# Create参数 0为执行命令行 5为fire a handler 6位发送email message 7为现实messagebox
$action = $TaskDefinition.Actions.Create(0)
$action.Path = "$taskCommand"
$action.Arguments = "$taskArgs"
# $action.WorkingDirectory

# 参数
# 1. taskName 
# 2. taskDefinition
# 3. flag (2:TASK_CREATE 4:TASK_UPDATE 6:TASK_CREATE_OR_UPDATE 8:TASK_DISABLE)
# 4. userid
# 5. password
# 6. logonType
#     (5: local system local service)
# 7. security descriptor
$rootFolder.RegisterTaskDefinition("$taskName",$taskDefinition,6,"System",$null,5)

$taskName = "newTask1"
$taskDescription =  "newTask1 Description"
$taskCommand = "powershell"
$taskScript = "test.ps1"
# $taskArgs = "-WindowStyle Hidden -NoInteractive -ExecutionPolicy unrestricted -file $taskScript"
$taskArgs = '-WindowStyle Hidden -ExecutionPolicy unrestricted -Command "& {echo hello>> E:\a.txt}"'
$taskStartTime = [datetime]::Now.AddMinutes(1)

# ����Schedule.Service��COM����
$service = new-object -ComObject("Schedule.Service")
$service.connect()

# ��ȡ���嶨ʱ�����Ŀ¼
$rootFolder = $service.GetFolder("\")

# ����һ����ʱ���񣬲���0�Ǳ���������������0
$taskDefinition = $service.NewTask(0)
# ���ö�ʱ��������ߡ������ʹ�������
$taskDefinition.RegistrationInfo.Description = "$taskDescrition"
$taskDefinition.RegistrationInfo.Author = "Administrator"

# ���ö�ʱ����(Enable�Ƿ����)������ʱ������
$taskDefinition.Settings.Enabled = $true
$taskDefinition.Settings.StartWhenAvailable = $true
$taskDefinition.Settings.Hidden = $false

# ���ô�����
# Create�� 1��ʾ�ƻ�ʱ 8��ʾ������ʱ
$trigger = $taskDefinition.Triggers.Create(1)
$trigger.Enabled = $true
$trigger.StartBoundary = $taskStartTime.ToString("yyyy-MM-dd'T'HH:mm:ss")

# duration��ʽ
# PnYnMnDTnHnMnS ���Ϊ1���ӣ�������Ϊ����
# $trigger.Repetition.Duration = 

# interval��ʽ
# PnDTnHnMnS �31�죬���1����
$trigger.Repetition.Interval = "PT1M"

# ���������ö�ʱ����Ĳ���
# Create���� 0Ϊִ�������� 5Ϊfire a handler 6λ����email message 7Ϊ��ʵmessagebox
$action = $TaskDefinition.Actions.Create(0)
$action.Path = "$taskCommand"
$action.Arguments = "$taskArgs"
# $action.WorkingDirectory

# ����
# 1. taskName 
# 2. taskDefinition
# 3. flag (2:TASK_CREATE 4:TASK_UPDATE 6:TASK_CREATE_OR_UPDATE 8:TASK_DISABLE)
# 4. userid
# 5. password
# 6. logonType
#     (5: local system local service)
# 7. security descriptor
$rootFolder.RegisterTaskDefinition("$taskName",$taskDefinition,6,"System",$null,5)

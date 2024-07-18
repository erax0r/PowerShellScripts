# Define your parameters
$switchIP = '192.168.1.2'
$username = 'admin'
$password = 'pw' | ConvertTo-SecureString -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password

# Set up the SSH session
$session = New-SSHSession -ComputerName $switchIP -Credential $credential
$shellstream = New-SSHShellStream -SessionId $session.SessionId

# Define the commands to run
$commands = @(
    'enable',
    'pw', # Replace with the actual password or a secure method to pass it
    'conf t',
    'interface gi1/0/2',
    'shut'
    'exit'
    'exit'
    'exit'
)

# Run the commands
foreach ($command in $commands) {
    Invoke-SSHStreamShellCommand -ShellStream $shellstream -Command $command
}

# Close the SSH session
Remove-SSHSession -SSHSession $session

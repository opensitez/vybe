# vybe-test: powershell/named_argument_binding/parameter_set_named
function Test { [CmdletBinding()] param([Parameter(ParameterSetName='A')]$a, [Parameter(ParameterSetName='B')]$b) if ($PSCmdlet.ParameterSetName -eq 'A') { $a } else { $b } }
if ((Test -a 1) -eq 1) { exit 0 }
exit 1

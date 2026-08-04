# vybe-test: powershell/type_inference/function_infer
function GetValue { 5 }
if ((GetValue) -is [int]) { exit 0 }
exit 1

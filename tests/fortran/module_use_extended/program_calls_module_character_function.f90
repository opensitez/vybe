! vybe-test: fortran/module_use_extended/program_calls_module_character_function
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module labels
implicit none
contains
function label_for(n) result(s)
integer, intent(in) :: n
character(len=4) :: s
if (n == 1) then
s = 'one'
else
s = 'many'
end if
end function label_for
end module labels
program t
use labels
if (trim(trim(label_for(1))) /= "one") then
    print *, "FAIL: want [one] got [", trim(label_for(1)), "]"
    stop 1
end if
end program t

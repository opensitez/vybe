! vybe-test: fortran/module_use_extended/use_only_character_name
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs
module names
implicit none
character(len=6) :: tag = 'alpha'
character(len=6) :: alt = 'beta'
end module names
program t
use names, only: tag
if (trim(trim(tag)) /= "alpha") then
    print *, "FAIL: want [alpha] got [", trim(tag), "]"
    stop 1
end if
end program t

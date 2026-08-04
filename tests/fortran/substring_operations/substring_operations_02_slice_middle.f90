! vybe-test: fortran/substring_operations/substring_operations_02_slice_middle
! origin: languages/fortran/tests/fortran/test_substring_operations.rs
program p
character(len=5) :: s='hello'
if (trim(trim(s(2:4))) /= "ell") then
    print *, "FAIL: want [ell] got [", trim(s(2:4)), "]"
    stop 1
end if
end program p

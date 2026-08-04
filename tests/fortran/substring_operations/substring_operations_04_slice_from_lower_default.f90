! vybe-test: fortran/substring_operations/substring_operations_04_slice_from_lower_default
! origin: languages/fortran/tests/fortran/test_substring_operations.rs
program p
character(len=5) :: s='hello'
if (trim(trim(s(3:))) /= "llo") then
    print *, "FAIL: want [llo] got [", trim(s(3:)), "]"
    stop 1
end if
end program p

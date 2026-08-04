! vybe-test: fortran/substring_operations/substring_operations_03_slice_to_upper_default
! origin: languages/fortran/tests/fortran/test_substring_operations.rs
program p
character(len=5) :: s='hello'
if (trim(trim(s(:3))) /= "hel") then
    print *, "FAIL: want [hel] got [", trim(s(:3)), "]"
    stop 1
end if
end program p

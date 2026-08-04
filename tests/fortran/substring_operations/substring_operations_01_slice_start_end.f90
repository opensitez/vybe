! vybe-test: fortran/substring_operations/substring_operations_01_slice_start_end
! origin: languages/fortran/tests/fortran/test_substring_operations.rs
program p
character(len=5) :: s='hello'
if (trim(trim(s(1:2))) /= "he") then
    print *, "FAIL: want [he] got [", trim(s(1:2)), "]"
    stop 1
end if
end program p

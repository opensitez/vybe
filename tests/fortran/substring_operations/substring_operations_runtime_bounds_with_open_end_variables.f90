! vybe-test: fortran/substring_operations/substring_operations_runtime_bounds_with_open_end_variables
! origin: languages/fortran/tests/fortran/test_substring_operations.rs

program p
character(len=6) :: s
integer :: start_idx
s = 'fortran'
start_idx = 3
if (trim(trim(s(start_idx:))) /= "tran") then
    print *, "FAIL: want [tran] got [", trim(s(start_idx:)), "]"
    stop 1
end if
if (trim(trim(s(1:start_idx))) /= "for") then
    print *, "FAIL: want [for] got [", trim(s(1:start_idx)), "]"
    stop 1
end if
end program p

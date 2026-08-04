! vybe-test: fortran/substring_operations/substring_operations_10_scan_into_slice
! origin: languages/fortran/tests/fortran/test_substring_operations.rs
program p
character(len=6) :: s='abcdef'
if ((scan(s(2:5), 'd')) /= 3) then
    print *, "FAIL: want [3] got [", scan(s(2:5), 'd'), "]"
    stop 1
end if
end program p

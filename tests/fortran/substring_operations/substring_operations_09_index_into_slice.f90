! vybe-test: fortran/substring_operations/substring_operations_09_index_into_slice
! origin: languages/fortran/tests/fortran/test_substring_operations.rs
program p
character(len=6) :: s='abcdef'
if ((index(s(2:), 'de')) /= 3) then
    print *, "FAIL: want [3] got [", index(s(2:), 'de'), "]"
    stop 1
end if
end program p

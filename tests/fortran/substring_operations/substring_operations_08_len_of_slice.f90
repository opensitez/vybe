! vybe-test: fortran/substring_operations/substring_operations_08_len_of_slice
! origin: languages/fortran/tests/fortran/test_substring_operations.rs
program p
character(len=6) :: s='abcdef'
if ((len(s(2:5))) /= 4) then
    print *, "FAIL: want [4] got [", len(s(2:5)), "]"
    stop 1
end if
end program p

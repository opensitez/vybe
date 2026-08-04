! vybe-test: fortran/substring_operations/substring_operations_07_single_char_index
! origin: languages/fortran/tests/fortran/test_substring_operations.rs
program p
character(len=6) :: s='abcdef'
if ((trim(s(6:6))) .neqv. .false.) then
    print *, "FAIL: want [f] got [", trim(s(6:6)), "]"
    stop 1
end if
end program p

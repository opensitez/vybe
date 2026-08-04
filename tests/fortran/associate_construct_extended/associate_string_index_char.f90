! vybe-test: fortran/associate_construct_extended/associate_string_index_char
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
character(len=5) :: s = 'abcde'
associate (ch => s(3:3))
if (trim(ch) /= "c") then
    print *, "FAIL: want [c] got [", ch, "]"
    stop 1
end if
end associate
end program t

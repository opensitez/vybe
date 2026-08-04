! vybe-test: fortran/io/print_char_array_join
! origin: languages/fortran/tests/fortran/test_io.rs

program test
    character(len=5) :: words(2)
    words(1) = "one"
    words(2) = "two"
    if (trim(words) /= "one two") then
    print *, "FAIL: want [one two] got [", words, "]"
    stop 1
end if
end program test

! vybe-test: fortran/print_io/print_multiple_lines
! origin: languages/fortran/tests/fortran/test_print_io.rs
program t
if (trim("a") /= "a") then
    print *, "FAIL: want [a] got [", "a", "]"
    stop 1
end if
if (trim("b") /= "b") then
    print *, "FAIL: want [b] got [", "b", "]"
    stop 1
end if
if (trim("c") /= "c") then
    print *, "FAIL: want [c] got [", "c", "]"
    stop 1
end if
end program t

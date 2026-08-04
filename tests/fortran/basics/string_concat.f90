! vybe-test: fortran/basics/string_concat
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    character(len=20) :: greeting
    greeting = "Hello" // " " // "World"
    if (trim(greeting) /= "Hello World") then
    print *, "FAIL: want [Hello World] got [", greeting, "]"
    stop 1
end if
end program test

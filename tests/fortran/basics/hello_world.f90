! vybe-test: fortran/basics/hello_world
! origin: languages/fortran/tests/fortran/test_basics.rs

program hello
    if (trim("Hello, World!") /= "Hello, World!") then
    print *, "FAIL: want [Hello, World!] got [", "Hello, World!", "]"
    stop 1
end if
end program hello

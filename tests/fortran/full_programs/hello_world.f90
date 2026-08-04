! vybe-test: fortran/full_programs/hello_world
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program hello
if (trim("Hello, World!") /= "Hello, World!") then
    print *, "FAIL: want [Hello, World!] got [", "Hello, World!", "]"
    stop 1
end if
end program hello

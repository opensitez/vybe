! vybe-test: fortran/modules_advanced/use_only_multiple
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module stuff
    integer :: a = 1, b = 2, c = 3
end module stuff

program test
    use stuff, only: a, c
    print *, a + c
end program test

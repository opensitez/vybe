! vybe-test: fortran/fortran2018/char_implied_type_concat
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    character(len=*), parameter :: greeting = 'Hello, ' // 'World!'
    print *, greeting
end program test

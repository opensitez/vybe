! vybe-test: fortran/fortran2018/co_sum_single_image
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: x = 42
    call co_sum(x)
    print *, x
end program test

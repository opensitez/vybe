! vybe-test: fortran/do_concurrent_extended/do_concurrent_shared
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a(5)
    integer :: factor
    factor = 3
    do concurrent (i = 1:5) shared(factor)
        a(i) = i * factor
    end do
    print *, a(2)
end program test

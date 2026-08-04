! vybe-test: fortran/fortran2008/do_concurrent_locality
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a(5), b(5)
    b = [1, 2, 3, 4, 5]
    do concurrent (i = 1:5) local(tmp)
        integer :: tmp
        tmp = b(i) * 2
        a(i) = tmp
    end do
    print *, a(3)
end program test

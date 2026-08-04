! vybe-test: fortran/programs/fibonacci
! origin: languages/fortran/tests/fortran/test_programs.rs

program fib
    integer :: n, a, b, temp, i
    n = 10
    a = 0
    b = 1
    do i = 1, n
        temp = a + b
        a = b
        b = temp
    end do
    print *, a
end program fib

! vybe-test: fortran/do_while/do_while_read_until_eof
! origin: languages/fortran/tests/fortran/test_do_while.rs

program test
    integer :: n, ios, s, u, i
    integer, dimension(4) :: nums
    nums = [1, 2, 3, 4]
    s = 0
    open(newunit=u, status='scratch', action='readwrite')
    do i = 1, 4
        write(u, '(I0)') nums(i)
    end do
    rewind(u)
    do while (.true.)
        read(u, *, iostat=ios) n
        if (ios /= 0) exit
        s = s + n
    end do
    print *, s
    close(u)
end program test

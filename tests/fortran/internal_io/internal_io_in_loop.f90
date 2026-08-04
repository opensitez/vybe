! vybe-test: fortran/internal_io/internal_io_in_loop
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=10) :: bufs(5)
    integer :: i
    do i = 1, 5
        write(bufs(i), '(I0)') i * i
    end do
    do i = 1, 5
        print *, trim(bufs(i))
    end do
end program test

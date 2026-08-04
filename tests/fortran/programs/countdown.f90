! vybe-test: fortran/programs/countdown
! origin: languages/fortran/tests/fortran/test_programs.rs

program countdown
    integer :: i
    do i = 10, 1, -1
        print *, i
    end do
    print *, "Launch!"
end program countdown

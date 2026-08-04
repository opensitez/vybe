! vybe-test: fortran/programs/simple_statistics
! origin: languages/fortran/tests/fortran/test_programs.rs

program stats
    integer :: i, n
    real :: sum, mean
    n = 5
    sum = 0.0
    do i = 1, n
        sum = sum + real(i)
    end do
    mean = sum / real(n)
    print *, "Mean:", mean
end program stats

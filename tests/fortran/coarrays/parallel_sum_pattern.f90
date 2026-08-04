! vybe-test: fortran/coarrays/parallel_sum_pattern
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x[*], total
    x = this_image()
    sync all
    if (this_image() == 1) then
        total = 0
        integer :: i
        do i = 1, num_images()
            total = total + x[i]
        end do
        print *, total
    end if
end program test

! vybe-test: fortran/named_loops/cycle_outer_skip_rest_of_inner
! origin: languages/fortran/tests/fortran/test_named_loops.rs

program test
    integer :: i, j
    outer: do i = 1, 3
        inner: do j = 1, 4
            if (j == 2) cycle outer
            print *, i, j
        end do inner
    end do outer
end program test

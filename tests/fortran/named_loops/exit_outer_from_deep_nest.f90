! vybe-test: fortran/named_loops/exit_outer_from_deep_nest
! origin: languages/fortran/tests/fortran/test_named_loops.rs

program test
    integer :: i, j, k
    outer: do i = 1, 10
        middle: do j = 1, 10
            inner: do k = 1, 10
                if (i + j + k == 10) exit outer
            end do inner
        end do middle
    end do outer
    print *, i, j, k
end program test

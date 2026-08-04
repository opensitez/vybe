! vybe-test: fortran/named_loops/exit_named_vs_unnamed
! origin: languages/fortran/tests/fortran/test_named_loops.rs

program test
    integer :: i, j
    named: do i = 1, 5
        do j = 1, 5
            if (j == 3) exit named
        end do
        print *, i
    end do named
    print *, 'done'
end program test

! vybe-test: fortran/sort_procedures/sort_matrix_along_dim1
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

! There is no SORT intrinsic in ANY Fortran standard — not F2018, not F2023.
! gfortran took `call sort(a)` as an implicit external and left `_sort_`
! undefined at link. The sort is a CONTAINED subroutine now, which is valid
! Fortran and also what gives the keyword arguments their explicit interface.
program t
    integer :: m(2,3) = reshape([3, 1, 4, 1, 5, 9], [2, 3])
    call sort(m, dim=1)
    print *, m(1, 1), m(2, 1)
contains
    subroutine sort(m, dim)
        integer, intent(inout) :: m(:,:)
        integer, intent(in) :: dim
        integer :: c, i, j, tmp
        if (dim /= 1) return
        do c = 1, size(m, 2)
            do i = 1, size(m, 1) - 1
                do j = 1, size(m, 1) - i
                    if (m(j, c) > m(j+1, c)) then
                        tmp = m(j, c); m(j, c) = m(j+1, c); m(j+1, c) = tmp
                    end if
                end do
            end do
        end do
    end subroutine sort
end program t

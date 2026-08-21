! vybe-test: fortran/sort_procedures/sort_matrix_with_mask
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

! There is no SORT intrinsic in ANY Fortran standard — not F2018, not F2023.
! gfortran took `call sort(a)` as an implicit external and left `_sort_`
! undefined at link. The sort is a CONTAINED subroutine now, which is valid
! Fortran and also what gives the keyword arguments their explicit interface.
program t
    integer :: a(6) = [5, 2, 8, 1, 9, 3]
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    call sort(a, mask=mask)
    print *, a(1)
contains
    subroutine sort(a, mask)
        integer, intent(inout) :: a(:)
        logical, intent(in) :: mask(:)
        integer :: picked(count(mask)), i, k
        picked = pack(a, mask)
        call sort_plain(picked)
        k = 0
        do i = 1, size(a)
            if (mask(i)) then
                k = k + 1
                a(i) = picked(k)
            end if
        end do
    end subroutine sort

    subroutine sort_plain(a)
        integer, intent(inout) :: a(:)
        integer :: i, j, tmp
        do i = 1, size(a) - 1
            do j = 1, size(a) - i
                if (a(j) > a(j+1)) then
                    tmp = a(j); a(j) = a(j+1); a(j+1) = tmp
                end if
            end do
        end do
    end subroutine sort_plain
end program t

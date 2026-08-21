! vybe-test: fortran/sort_procedures/sort_integer_vector_descending_runtime
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

! There is no SORT intrinsic in ANY Fortran standard — not F2018, not F2023.
! gfortran took `call sort(a)` as an implicit external and left `_sort_`
! undefined at link. The sort is a CONTAINED subroutine now, which is valid
! Fortran and also what gives the keyword arguments their explicit interface.
program t
    integer :: a(4) = [3, 1, 4, 2]
    call sort(a, reverse=.true.)
    if ((a(1)) /= 4) then
    print *, "FAIL: want [4] got [", a(1), "]"
    stop 1
end if
    if ((a(2)) /= 3) then
    print *, "FAIL: want [3] got [", a(2), "]"
    stop 1
end if
    if ((a(3)) /= 2) then
    print *, "FAIL: want [2] got [", a(3), "]"
    stop 1
end if
    if ((a(4)) /= 1) then
    print *, "FAIL: want [1] got [", a(4), "]"
    stop 1
end if
contains
    subroutine sort(a, reverse)
        integer, intent(inout) :: a(:)
        logical, intent(in), optional :: reverse
        integer :: i, j, tmp
        logical :: down
        down = .false.
        if (present(reverse)) down = reverse
        do i = 1, size(a) - 1
            do j = 1, size(a) - i
                if ((.not. down .and. a(j) > a(j+1)) .or. &
                    (down .and. a(j) < a(j+1))) then
                    tmp = a(j); a(j) = a(j+1); a(j+1) = tmp
                end if
            end do
        end do
    end subroutine sort
end program t

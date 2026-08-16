! vybe-test: fortran/arrays/type_bound_subroutine_populates_allocatable_array
! origin: languages/fortran/tests/fortran/test_arrays.rs
module m
    type :: list
    contains
        procedure :: fill
    end type list
contains
    subroutine fill(self, arr)
        class(list), intent(in) :: self
        integer, allocatable, intent(out) :: arr(:)
        allocate(arr(3))
        arr(1) = 5
        arr(2) = 3
        arr(3) = 8
    end subroutine fill
end module m
program driver
use m
    type(list) :: value
    integer, allocatable :: arr(:)

    call value%fill(arr)
    if ((arr(1)) /= 5) then
    print *, "FAIL: want [5] got [", arr(1), "]"
    stop 1
end if
    if ((arr(2)) /= 3) then
    print *, "FAIL: want [3] got [", arr(2), "]"
    stop 1
end if
    if ((arr(3)) /= 8) then
    print *, "FAIL: want [8] got [", arr(3), "]"
    stop 1
end if

end program driver

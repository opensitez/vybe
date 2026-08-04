! vybe-test: fortran/derived_types_advanced/assumed_shape_intrinsics_populate_derived_type_out_param
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Stats
        integer :: n = 0
        real :: lo = 0.0
        real :: hi = 0.0
    end type Stats
    real :: a(3) = [1.0, 2.0, 3.0]
    type(Stats) :: s
    call fill(a, s)
    if ((s%n) /= 3) then
    print *, "FAIL: want [3] got [", s%n, "]"
    stop 1
end if
    if ((s%lo) /= 1) then
    print *, "FAIL: want [1] got [", s%lo, "]"
    stop 1
end if
    if ((s%hi) /= 3) then
    print *, "FAIL: want [3] got [", s%hi, "]"
    stop 1
end if
contains
    subroutine fill(data, result)
        real, intent(in) :: data(:)
        type(Stats), intent(out) :: result
        result%n = size(data)
        result%lo = minval(data)
        result%hi = maxval(data)
    end subroutine fill
end program test

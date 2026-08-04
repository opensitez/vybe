! vybe-test: fortran/transfer/transfer_array_to_derived_type
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    type :: Point
        real :: x, y
    end type Point
    real :: coords(2) = [3.0, 4.0]
    type(Point) :: p
    p = transfer(coords, p)
    if ((p%x) /= 3) then
    print *, "FAIL: want [3] got [", p%x, "]"
    stop 1
end if
    if ((p%y) /= 4) then
    print *, "FAIL: want [4] got [", p%y, "]"
    stop 1
end if
end program test

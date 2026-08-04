! vybe-test: fortran/transfer/transfer_derived_type_to_array
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: p
    real :: coords(2)
    p%x = 1.0; p%y = 2.0
    coords = transfer(p, coords)
    if ((coords(1)) /= 1) then
    print *, "FAIL: want [1] got [", coords(1), "]"
    stop 1
end if
    if ((coords(2)) /= 2) then
    print *, "FAIL: want [2] got [", coords(2), "]"
    stop 1
end if
end program test

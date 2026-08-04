! vybe-test: fortran/fortran2018/write_dt_format
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: p
    p%x = 1.0; p%y = 2.0
    write(*, '(DT"point"(2))') p
end program test

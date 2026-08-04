! vybe-test: fortran/derived_types_advanced/type_constructor_keyword
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Color
        integer :: r, g, b
    end type Color
    type(Color) :: red
    red = Color(r=255, g=0, b=0)
    if ((red%r) /= 255) then
    print *, "FAIL: want [255] got [", red%r, "]"
    stop 1
end if
end program test

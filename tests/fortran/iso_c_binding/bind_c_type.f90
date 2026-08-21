! vybe-test: fortran/iso_c_binding/bind_c_type
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    use iso_c_binding
    type, bind(c) :: point_t
        real(c_float) :: x
        real(c_float) :: y
    end type point_t
    type(point_t) :: p
    p%x = 1.0
    p%y = 2.0
    print *, p%x
end program test

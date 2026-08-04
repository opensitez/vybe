! vybe-test: fortran/fortran2003/c_interop_kinds
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    use iso_c_binding
    integer(c_int)    :: i = 1_c_int
    integer(c_long)   :: j = 2_c_long
    integer(c_size_t) :: k = 3_c_size_t
    real(c_float)     :: f = 1.0_c_float
    real(c_double)    :: d = 2.0_c_double
    logical(c_bool)   :: b = .true._c_bool
    character(len=1, kind=c_char) :: ch = c_null_char
    print *, i, j, f
end program test

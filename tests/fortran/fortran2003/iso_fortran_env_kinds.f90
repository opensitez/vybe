! vybe-test: fortran/fortran2003/iso_fortran_env_kinds
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    use iso_fortran_env
    integer(int32) :: n = 42_int32
    integer(int64) :: big = 1000000000_int64
    real(real32) :: f = 3.14_real32
    real(real64) :: d = 3.14159265_real64
    print *, n
    print *, big
end program test

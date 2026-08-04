! vybe-test: fortran/explicit_interfaces/explicit_interfaces_03
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
interface
real function f(x)
real::x
end function f
end interface

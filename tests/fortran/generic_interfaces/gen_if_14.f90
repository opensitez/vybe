! vybe-test: fortran/generic_interfaces/gen_if_14
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface g
module procedure i1,i2
end interface
contains
integer function i1()
i1=1
end
integer function i2()
i2=2
end
end module m

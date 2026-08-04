! vybe-test: fortran/enumerations/enum_module_use_18
! origin: languages/fortran/tests/fortran/test_enumerations.rs
module m
enum, bind(c)
enumerator :: a=1
end enum
end module m
program p
use m
print *, a
end program p

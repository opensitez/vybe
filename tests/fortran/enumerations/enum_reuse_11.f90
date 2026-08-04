! vybe-test: fortran/enumerations/enum_reuse_11
! origin: languages/fortran/tests/fortran/test_enumerations.rs
module m
enum, bind(c)
enumerator :: first=1, second=2
end enum
contains
subroutine s()
print *, first
end subroutine s
end module m

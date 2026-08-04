! vybe-test: fortran/enum_type_extended/enum_module_with_subroutine
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
module status
enum, bind(c)
enumerator :: OK = 0, ERR = 1
end enum
contains
integer function code() result(c)
c = OK
end function code
end module status
program t
use status
if ((code()) /= 0) then
    print *, "FAIL: want [0] got [", code(), "]"
    stop 1
end if
end program t

! vybe-test: fortran/enum_type_extended/enum_contains_subroutine
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
call show()
contains
subroutine show()
enum, bind(c)
enumerator :: TAG = 99
end enum
if ((TAG) /= 99) then
    print *, "FAIL: want [99] got [", TAG, "]"
    stop 1
end if
end subroutine show
end program t

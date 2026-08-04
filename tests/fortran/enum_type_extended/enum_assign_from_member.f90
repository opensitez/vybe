! vybe-test: fortran/enum_type_extended/enum_assign_from_member
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 5, B = 10
end enum
integer :: x
x = A
if ((x) /= 5) then
    print *, "FAIL: want [5] got [", x, "]"
    stop 1
end if
x = B
if ((x) /= 10) then
    print *, "FAIL: want [10] got [", x, "]"
    stop 1
end if
end program t

! vybe-test: fortran/enum_type_extended/enum_nested_select_in_function
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
if ((label(2)) /= 20) then
    print *, "FAIL: want [20] got [", label(2), "]"
    stop 1
end if
contains
integer function label(v) result(r)
enum, bind(c)
enumerator :: ONE = 1, TWO = 2
end enum
select case (v)
case (ONE)
r = 10
case (TWO)
r = 20
case default
r = 0
end select
end function label
end program t

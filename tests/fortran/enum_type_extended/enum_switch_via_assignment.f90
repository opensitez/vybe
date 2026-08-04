! vybe-test: fortran/enum_type_extended/enum_switch_via_assignment
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: MODE_A = 1, MODE_B = 2
end enum
integer :: mode
mode = MODE_A
if ((mode) /= 1) then
    print *, "FAIL: want [1] got [", mode, "]"
    stop 1
end if
mode = MODE_B
if ((mode) /= 2) then
    print *, "FAIL: want [2] got [", mode, "]"
    stop 1
end if
end program t

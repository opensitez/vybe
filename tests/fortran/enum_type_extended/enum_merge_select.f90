! vybe-test: fortran/enum_type_extended/enum_merge_select
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: CHOICE_A = 10, CHOICE_B = 20
end enum
if ((merge(CHOICE_A, CHOICE_B, .true.)) /= 10) then
    print *, "FAIL: want [10] got [", merge(CHOICE_A, CHOICE_B, .true.), "]"
    stop 1
end if
end program t

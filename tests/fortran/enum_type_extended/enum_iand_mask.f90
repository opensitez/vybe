! vybe-test: fortran/enum_type_extended/enum_iand_mask
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: MASK = 7, VAL = 5
end enum
if ((iand(MASK, VAL)) /= 5) then
    print *, "FAIL: want [5] got [", iand(MASK, VAL), "]"
    stop 1
end if
end program t

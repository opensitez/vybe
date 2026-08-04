! vybe-test: fortran/enum_type_extended/enum_array_lookup
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: RED = 0, GREEN = 1, BLUE = 2
end enum
character(len=5) :: names(0:2)
names(RED) = 'red'
names(GREEN) = 'grn'
names(BLUE) = 'blu'
if (trim(trim(names(GREEN))) /= "grn") then
    print *, "FAIL: want [grn] got [", trim(names(GREEN)), "]"
    stop 1
end if
end program t

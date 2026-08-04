! vybe-test: fortran/enum_type_extended/enum_member_as_array_size
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: SIZE = 4
end enum
integer :: arr(SIZE)
arr = [1, 2, 3, 4]
if ((arr(SIZE)) /= 4) then
    print *, "FAIL: want [4] got [", arr(SIZE), "]"
    stop 1
end if
end program t

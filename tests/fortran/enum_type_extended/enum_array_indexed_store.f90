! vybe-test: fortran/enum_type_extended/enum_array_indexed_store
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 1, B = 2, C = 3
end enum
integer :: arr(3)
arr(A) = 10
arr(B) = 20
arr(C) = 30
if ((arr(B)) /= 20) then
    print *, "FAIL: want [20] got [", arr(B), "]"
    stop 1
end if
end program t

! vybe-test: fortran/enum_type_extended/enum_array_sum_members
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 1, B = 2, C = 3, D = 4
end enum
integer :: vals(4) = [A, B, C, D]
if ((sum(vals)) /= 10) then
    print *, "FAIL: want [10] got [", sum(vals), "]"
    stop 1
end if
end program t

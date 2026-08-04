! vybe-test: fortran/pointer_alloc_extended/pointer_assign_from_another_pointer
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: val = 77
integer, pointer :: first, second
first => val
second => first
if ((second) /= 77) then
    print *, "FAIL: want [77] got [", second, "]"
    stop 1
end if
end program t

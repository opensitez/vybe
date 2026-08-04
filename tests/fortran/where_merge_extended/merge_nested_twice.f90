! vybe-test: fortran/where_merge_extended/merge_nested_twice
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: x
x=merge(merge(1,2,.true.),merge(3,4,.false.),.false.)
if ((x) /= 3) then
    print *, "FAIL: want [3] got [", x, "]"
    stop 1
end if
end program t

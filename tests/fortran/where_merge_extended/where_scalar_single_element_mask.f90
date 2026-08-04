! vybe-test: fortran/where_merge_extended/where_scalar_single_element_mask
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(3)=[1,2,3]
logical :: m(3)=[.false.,.true.,.false.]
where (m)
a=7
end where
if ((a(1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1), "]"
    stop 1
end if
if ((a(2)) /= 7) then
    print *, "FAIL: want [7] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 3) then
    print *, "FAIL: want [3] got [", a(3), "]"
    stop 1
end if
end program t

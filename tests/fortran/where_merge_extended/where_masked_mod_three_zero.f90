! vybe-test: fortran/where_merge_extended/where_masked_mod_three_zero
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(6)=[1,2,3,4,5,6]
where (mod(a,3)==0)
a=0
end where
if ((a(3)) /= 0) then
    print *, "FAIL: want [0] got [", a(3), "]"
    stop 1
end if
if ((a(4)) /= 4) then
    print *, "FAIL: want [4] got [", a(4), "]"
    stop 1
end if
if ((a(6)) /= 0) then
    print *, "FAIL: want [0] got [", a(6), "]"
    stop 1
end if
end program t

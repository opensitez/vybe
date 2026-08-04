! vybe-test: fortran/where_merge_extended/where_masked_real_neg_to_zero
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
real :: a(4)=[1.5,-2.0,3.0,-0.5]
where (a<0.0)
a=0.0
end where
if ((a(2)) /= 0) then
    print *, "FAIL: want [0] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 3) then
    print *, "FAIL: want [3] got [", a(3), "]"
    stop 1
end if
end program t

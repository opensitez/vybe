! vybe-test: fortran/where_merge_extended/where_precomputed_logical_mask
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[1,2,3,4]
logical :: m(4)=[.true.,.false.,.true.,.false.]
integer :: b(4)=0
where (m)
b=a*10
end where
if ((b(1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 0) then
    print *, "FAIL: want [0] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 30) then
    print *, "FAIL: want [30] got [", b(3), "]"
    stop 1
end if
end program t

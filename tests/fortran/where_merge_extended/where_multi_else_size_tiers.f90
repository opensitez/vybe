! vybe-test: fortran/where_merge_extended/where_multi_else_size_tiers
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[3,15,50,200]
integer :: t(4)
where (a<10)
t=1
elsewhere (a<100)
t=2
elsewhere
t=3
end where
if ((t(1)) /= 1) then
    print *, "FAIL: want [1] got [", t(1), "]"
    stop 1
end if
if ((t(2)) /= 2) then
    print *, "FAIL: want [2] got [", t(2), "]"
    stop 1
end if
if ((t(4)) /= 3) then
    print *, "FAIL: want [3] got [", t(4), "]"
    stop 1
end if
end program t

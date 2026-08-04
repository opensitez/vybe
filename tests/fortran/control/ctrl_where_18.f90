! vybe-test: fortran/control/ctrl_where_18
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer::a(3)=[1,2,3]
where(a>1) a=a+1
if ((a(1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1), "]"
    stop 1
end if
if ((a(2)) /= 3) then
    print *, "FAIL: want [3] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 4) then
    print *, "FAIL: want [4] got [", a(3), "]"
    stop 1
end if
end program p

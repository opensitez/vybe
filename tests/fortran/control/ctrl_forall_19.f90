! vybe-test: fortran/control/ctrl_forall_19
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer::a(3), i
forall(i=1:3) a(i)=i*2
if ((a(1)) /= 2) then
    print *, "FAIL: want [2] got [", a(1), "]"
    stop 1
end if
if ((a(2)) /= 4) then
    print *, "FAIL: want [4] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 6) then
    print *, "FAIL: want [6] got [", a(3), "]"
    stop 1
end if
end program p

! vybe-test: fortran/control/ctrl_where_elsewhere_24
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer::a(4)
integer::b(4)
a = (/1, 2, 3, 4/)
b = 0
where (a <= 2)
  b = 10
elsewhere (a <= 3)
  b = 20
elsewhere
  b = 30
end where
if ((b(1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 10) then
    print *, "FAIL: want [10] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 20) then
    print *, "FAIL: want [20] got [", b(3), "]"
    stop 1
end if
if ((b(4)) /= 30) then
    print *, "FAIL: want [30] got [", b(4), "]"
    stop 1
end if
end program p

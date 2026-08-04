! vybe-test: fortran/control_flow_extended/where_elsewhere_masked_update
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: a(5), b(5)
a = 0
b = 1
where (mod(a, 2) == 0)
  b = b + 10
elsewhere (a < 0)
  b = 99
elsewhere
  b = 7
end where
if ((b(1)) /= 11) then
    print *, "FAIL: want [11] got [", b(1), "]"
    stop 1
end if
if ((b(3)) /= 11) then
    print *, "FAIL: want [11] got [", b(3), "]"
    stop 1
end if
if ((b(5)) /= 11) then
    print *, "FAIL: want [11] got [", b(5), "]"
    stop 1
end if
end program t

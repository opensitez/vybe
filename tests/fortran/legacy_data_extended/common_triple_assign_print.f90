! vybe-test: fortran/legacy_data_extended/common_triple_assign_print
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: p, q, r
common /trio/ p, q, r
p = 11; q = 22; r = 33
if ((p) /= 11) then
    print *, "FAIL: want [11] got [", p, "]"
    stop 1
end if
if ((q) /= 22) then
    print *, "FAIL: want [22] got [", q, "]"
    stop 1
end if
if ((r) /= 33) then
    print *, "FAIL: want [33] got [", r, "]"
    stop 1
end if
end program t

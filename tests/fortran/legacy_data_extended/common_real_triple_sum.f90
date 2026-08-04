! vybe-test: fortran/legacy_data_extended/common_real_triple_sum
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
real :: u, v, w
common /rv/ u, v, w
u = 1.0; v = 2.0; w = 3.0
if ((u + v + w) /= 6) then
    print *, "FAIL: want [6] got [", u + v + w, "]"
    stop 1
end if
end program t

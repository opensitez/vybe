! vybe-test: fortran/select_type_rank_extended/select_type_real_fractional
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(real :: val)
val = 0.25
select type(val)
type is (real)
if ((int(val * 100.0)) /= 25) then
    print *, "FAIL: want [25] got [", int(val * 100.0), "]"
    stop 1
end if
end select
end program t

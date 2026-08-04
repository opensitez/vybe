! vybe-test: fortran/select_type_rank_extended/select_type_real_negative
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(real :: val)
val = -4.0
select type(val)
type is (real)
if ((int(abs(val))) /= 4) then
    print *, "FAIL: want [4] got [", int(abs(val)), "]"
    stop 1
end if
end select
end program t

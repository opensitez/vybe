! vybe-test: fortran/select_type_rank_extended/select_type_reallocate_change_type
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(integer :: val)
val = 8
select type(val)
type is (integer)
if ((val) /= 8) then
    print *, "FAIL: want [8] got [", val, "]"
    stop 1
end if
end select
deallocate(val)
allocate(real :: val)
val = 2.5
select type(val)
type is (real)
if ((int(val)) /= 2) then
    print *, "FAIL: want [2] got [", int(val), "]"
    stop 1
end if
end select
end program t

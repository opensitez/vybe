! vybe-test: fortran/select_type_rank_extended/select_type_integer_zero
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(integer :: val)
val = 0
select type(val)
type is (integer)
if ((val) /= 0) then
    print *, "FAIL: want [0] got [", val, "]"
    stop 1
end if
end select
end program t

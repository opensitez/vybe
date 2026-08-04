! vybe-test: fortran/select_type_rank_extended/select_type_integer_modify_in_branch
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(integer :: val)
val = 10
select type(val)
type is (integer)
val = val + 5
if ((val) /= 15) then
    print *, "FAIL: want [15] got [", val, "]"
    stop 1
end if
end select
end program t

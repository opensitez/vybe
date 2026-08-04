! vybe-test: fortran/select_type_rank_extended/select_type_integer_large
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(integer :: val)
val = 1000000
select type(val)
type is (integer)
if ((val / 1000000) /= 1) then
    print *, "FAIL: want [1] got [", val / 1000000, "]"
    stop 1
end if
end select
end program t

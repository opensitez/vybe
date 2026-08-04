! vybe-test: fortran/select_type_rank_extended/select_type_logical_false
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(logical :: val)
val = .false.
select type(val)
type is (logical)
if ((val) .neqv. .false.) then
    print *, "FAIL: want [false] got [", val, "]"
    stop 1
end if
end select
end program t

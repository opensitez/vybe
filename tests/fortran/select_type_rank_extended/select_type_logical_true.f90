! vybe-test: fortran/select_type_rank_extended/select_type_logical_true
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(logical :: val)
val = .true.
select type(val)
type is (logical)
if ((val) .neqv. .true.) then
    print *, "FAIL: want [true] got [", val, "]"
    stop 1
end if
end select
end program t

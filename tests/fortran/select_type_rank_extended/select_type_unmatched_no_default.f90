! vybe-test: fortran/select_type_rank_extended/select_type_unmatched_no_default
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(real :: val)
val = 1.5
select type(val)
type is (integer)
print *, 7
type is (logical)
print *, 8
end select
end program t

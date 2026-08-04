! vybe-test: fortran/select_type_rank_extended/select_type_class_default_only
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(integer :: val)
val = 55
select type(val)
class default
if ((val) /= 55) then
    print *, "FAIL: want [55] got [", val, "]"
    stop 1
end if
end select
end program t

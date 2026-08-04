! vybe-test: fortran/select_type_rank_extended/select_rank_logical_vector_any
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
call tag([.true., .false., .true.])
contains
subroutine tag(x)
logical, intent(in) :: x(..)
select rank(x)
rank(1)
if (any(x)) print *, 1
rank default
if ((0) /= 1) then
    print *, "FAIL: want [1] got [", 0, "]"
    stop 1
end if
end select
end subroutine tag
end program t

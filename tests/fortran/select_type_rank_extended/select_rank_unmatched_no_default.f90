! vybe-test: fortran/select_type_rank_extended/select_rank_unmatched_no_default
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
call tag([1, 2, 3])
contains
subroutine tag(x)
integer, intent(in) :: x(..)
select rank(x)
rank(0)
print *, x
end select
end subroutine tag
end program t

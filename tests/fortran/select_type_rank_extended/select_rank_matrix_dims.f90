! vybe-test: fortran/select_type_rank_extended/select_rank_matrix_dims
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
call tag(reshape([1,2,3,4,5,6], [2,3]))
contains
subroutine tag(x)
integer, intent(in) :: x(..)
select rank(x)
rank(2)
print *, size(x,1), size(x,2)
rank default
print *, 0, 0
end select
end subroutine tag
end program t

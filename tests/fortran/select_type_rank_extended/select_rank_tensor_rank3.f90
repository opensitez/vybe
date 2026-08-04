! vybe-test: fortran/select_type_rank_extended/select_rank_tensor_rank3
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
call tag(reshape([(i, i=1,8)], [2,2,2]))
contains
subroutine tag(x)
integer, intent(in) :: x(..)
select rank(x)
rank(3)
print *, size(x,1), size(x,2), size(x,3)
rank default
print *, 0, 0, 0
end select
end subroutine tag
end program t

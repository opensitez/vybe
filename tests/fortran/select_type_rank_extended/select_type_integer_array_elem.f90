! vybe-test: fortran/select_type_rank_extended/select_type_integer_array_elem
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: arr(:)
allocate(integer :: arr(3))
arr = [4, 5, 6]
select type(arr(2))
type is (integer)
if ((arr(2)) /= 5) then
    print *, "FAIL: want [5] got [", arr(2), "]"
    stop 1
end if
end select
end program t

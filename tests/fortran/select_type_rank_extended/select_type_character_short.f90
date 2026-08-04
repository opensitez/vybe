! vybe-test: fortran/select_type_rank_extended/select_type_character_short
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(character(len=3) :: val)
val = 'abc'
select type(val)
type is (character(len=*))
if ((len_trim(val)) /= 3) then
    print *, "FAIL: want [3] got [", len_trim(val), "]"
    stop 1
end if
end select
end program t

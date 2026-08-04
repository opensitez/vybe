! vybe-test: fortran/select_type_rank_extended/select_type_character_empty
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
class(*), allocatable :: val
allocate(character(len=1) :: val)
val = 'x'
select type(val)
type is (character(len=*))
if (trim(val) /= "x") then
    print *, "FAIL: want [x] got [", val, "]"
    stop 1
end if
end select
end program t

! vybe-test: fortran/select_type_rank_extended/select_rank_assumed_rank_function_result
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
if ((pick(reshape([1,2,3,4],[2,2]))) /= 5) then
    print *, "FAIL: want [5] got [", pick(reshape([1,2,3,4],[2,2])), "]"
    stop 1
end if
contains
integer function pick(m)
integer, intent(in) :: m(..)
select rank(m)
rank(2)
pick = m(1,1) + m(2,2)
rank default
pick = 0
end select
end function pick
end program t

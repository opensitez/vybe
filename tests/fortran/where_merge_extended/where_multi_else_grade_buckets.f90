! vybe-test: fortran/where_merge_extended/where_multi_else_grade_buckets
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: p(4)=[55,72,88,95]
character(len=1) :: g(4)
where (p<60)
g="F"
elsewhere (p<70)
g="D"
elsewhere (p<80)
g="C"
elsewhere
g="A"
end where
if ((g(1)) .neqv. .false.) then
    print *, "FAIL: want [F] got [", g(1), "]"
    stop 1
end if
if (trim(g(2)) /= "C") then
    print *, "FAIL: want [C] got [", g(2), "]"
    stop 1
end if
if (trim(g(4)) /= "A") then
    print *, "FAIL: want [A] got [", g(4), "]"
    stop 1
end if
end program t

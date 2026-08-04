! vybe-test: fortran/strings_extended/count_occurrences_runtime
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
if ((count_occurrences('the quick brown fox jumps over the lazy dog', 'the')) /= 2) then
    print *, "FAIL: want [2] got [", count_occurrences('the quick brown fox jumps over the lazy dog', 'the'), "]"
    stop 1
end if
contains
pure function count_occurrences(s, sub) result(n)
character(len=*), intent(in) :: s, sub
integer :: n, pos, start, lsub
n = 0
lsub = len_trim(sub)
if (lsub == 0) return
start = 1
do
    pos = index(s(start:), trim(sub))
    if (pos == 0) exit
    n = n + 1
    start = start + pos + lsub - 1
end do
end function count_occurrences
end program t

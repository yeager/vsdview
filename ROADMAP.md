# Proposed next improvements

Prioritized follow-up proposals; these are not implemented in this release.

1. Convert large documents in a worker with cancellation and progress; commit window state only after a successful result.
2. Add lazy page rendering and a bounded page cache to avoid converting the whole document before first display.
3. Preserve selected page, zoom and pan per recent document.
4. Add interactive GTK smoke tests for file opening, page navigation, export and window close on Linux and Windows.

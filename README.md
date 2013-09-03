SolutionExplorerPlus
====================

Adding useful functionality to solution explorer - plugin for VS2012.

Features
--------

- Repopulate

Provides a simple way to add any missing files of similar types to your existing project filters. Can be run on a single item to add all files of the same extension, in the same folder, to the same filter. Can be run at solution top-level to perform recursively over entire solution.

Known Bugs
----------

- Doesn't handle files added directly to a project node.
- Project recursive add doesn't work properly.
- Doesn't handle toggling on and off add-in at runtime (must re-open solution).



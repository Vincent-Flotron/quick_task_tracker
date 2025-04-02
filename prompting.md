Write me a python app using tkinter and sqlite that allow me to create, edit, suppress Tasks.
The window will display a table with sortable columns in a kind of treeview because **task** can have sub **tasks**.
And below, some fields described at the last lines.

The tables relations will be:

```
task (one)
  task (0 or many)
  link (0 or many)
  delivery (0 or many)
    tag
  tag (0 or many)
  origin (one)
```

And the table config will be:
```
task
  id
  customer
  name
  description 
  started_at
  finished_at
  task_id
```

```
delivery
  id
  version
  server
  delivery_date_time
```

```
task_delivery
  id
  delivery_id
  task_id
```

```
link
  id
  type
  raw_link
```

```
task_link
  id
  task_id
  link_id
```

```
tag_link
  id
  link_id
  tag_id
```

```
tag_task
  id
  task_id
  tag_id
```

```
tag
  id
  type
  keywords
```

```
origin
  id
  type
  raw_link
```

```
task_origin
  id
  task_id
  origin_id
```

```
booking
  id
  description
  started_at (date_time)
  ended_at (date_time)
  duration (time)
  task_id
  origin_id
``` 


The fields displayed into the sortable table must be:
task.
  customer
  name
  description 
  started_at
  finished_at

The displayed fields must be the same as the ones displayed in the sortable table and all the tables related to the task selected in the sortable table.

======

Great. Could you add possibility to book?
A booking must be linked to:
- a task and
- an origin

So, when I click on a task and, the bookings linked to this task and its origin must be displayed on a panel on the right.

And to make the origins re-usable, when adding an new origin to a task, make possible to select an existing origin. Show available origins using a combobox that display matching origin when entering their name. So by the way, consider the origin table already get the field "name"

Implement it following the description of this new table below:
```
booking
  id
  description
  started_at (date_time)
  ended_at (date_time)
  duration (time)
  task_id
  origin_id
```


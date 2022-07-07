# Start Form Filling

## Introduction

Look at the new Start template and work through how to make filling the white spaces easier.
The salmon colored spaces will be pulled from EoS, the white spaces will need to be filled in.

### Start Input Form

There are the fields that would need to be filled:

- Job # : Can be pulled from the Database
- Subjob # : User Input when creating the Start?
- Plan : Can be split form the Selection.
- Elv/Opt : Can be split from the Selection.
- Contract Amt: Should be pulled from the database for the corresponding plan.
- Contract #: Should be pulled from the database for the corresponding executed contract.
- Contract File Name: Should be pulled from the database.
- Estimate Name: Should be pulled from the database.

## Thought Process

Overall, here is my current approach to this.

We will pull the job from the database. I'm thinking that this would be a pk of the job. The Subjob # would be one of the pk's of the subjob. I'm thinking it would be (Job # - Subjob #). 

For the plans, If the options are all selected in before the start is generated, then we will need to split the selection into Plan, Elv/Opt, Elv/Opt Desc, and Contract Amt. We can check if the ELV is set to all, and if it is, have a prompt to change it to the correct ELV. The contract amount would correspond with the plan.

I'm assuming that The contract # and filename are associated with the latest executed contract. In theory, should be in a database. Along with the estimate name. If they are, we can connect to the database and pull the latest executed contract and the latest estimate. We can then compare the latest estimate values to the latest executed contract values. If they are different, we can do an "Add For currect Estimate" and show the difference in values in the form.

Since we are using Microsoft SQL for the database and excel for the rest, it might make sense to switch to C#. Microsoft designed C#, and it has a bunch of built it features for connecting to Excel and SQL.


## Issues

At the current moment, I do not have access to the Database.

I noticed that the Lot and Address information is populated through EoS. Will this information be filled out ahead of time?

N
# Database Design — Explanation
 
This section explains the reasoning behind the placement of each field in the database.  
The goal of the structure is to keep the design **normalized, logical, and easily maintainable**.
 
---
 
### `Login` table
Holds individual login events.
 
- `id`: unique identifier for each login event.  
- `date`, `time`: separated for easier filtering and reporting.  
- `pc_id`: identifies which PC the login happened on.  
- `user_id`: identifies the user who logged in.  
- `freediskspace`: stored here because the free space can change every time someone logs in — it’s a **time-dependent** value.  
  Keeping it here preserves historical values for every login event.
 
---
 
### `User` table
Contains basic user information.
 
- `id`: primary key for unique identification.  
- `name`: username, stored once to avoid duplication across logins.
 
---
 
### `PC` table
Represents the physical or virtual machine.
 
- `id`: primary key.  
- `name`: device or hostname.  
- `deviceID`: references what type of device it is (from the `Device` table).  
- `modelID`: connects to the model information (in `Model`).  
- `ram`: amount of memory in GB.  
- `processorID`: links to processor details.  
- `operatingsystemID`: links to the installed operating system.  
- `operatingsysteminstallationdate`: specific to this PC, so it belongs here.  
- `disk`: stores the drive letter (e.g., `C:`). The **actual free space** is not here because it changes over time — that’s tracked in `Login`.  
- `note`: any additional info about the PC (comments, remarks, etc.).
 
---
 
### `Device` table
Defines the **type** of device.
 
- `id`: primary key.  
- `type`: category such as `Laptop`, `Desktop`, `Server`, etc.  
  The `PC` table references this instead of repeating text for each row.
 
---
 
### `Brand` table
Stores manufacturer names.
 
- `id`: primary key.  
- `name`: e.g., `Dell`, `HP`, `Lenovo`.  
  It’s separated so models can reference it and avoid repeating brand names.
 
---
 
### `Model` table
Defines the model of the PC.
 
- `id`: primary key.  
- `brandID`: connects the model to a brand.  
- `name`: the specific model name.  
  This structure prevents data duplication and allows multiple models per brand.
 
---
 
### `Processor` table
Stores processor instances or codes.
 
- `id`: primary key.  
- `processormodelID`: connects to the processor model table.  
- `processorCode`: a unique identifier or code for the CPU (e.g., `i7-9700K`).  
  Separating processor model and processor code allows flexibility for different versions or batches.
 
---
 
### `ProcessorModel` table
Contains general CPU model names.
 
- `id`: primary key.  
- `name`: the processor family/model name (e.g., `Intel Core i7`).  
  Used to group processors that share the same model type.
 
---
 
### `OperatingSystem` table
Defines available operating systems.
 
- `id`: primary key.  
- `name`: e.g., `Windows 10`, `Windows 11 Pro`, etc.  
  Only OS types are stored here; installation dates are stored in `PC`, because they vary per machine.
 
---
 
### Summary of Key Design Decisions
 
- **Free disk space** → in `Login`, because it changes over time.  
- **OS installation date** → in `PC`, because it’s specific to the machine, not to the OS type.  
- **Brand–Model split** → avoids repeating brand names.  
- **Processor–ProcessorModel split** → supports multiple variants under one CPU model.  
- **Device table** → helps categorize PCs without text duplication.  
- **Normalization** → each attribute belongs only where it logically depends on the key (3NF).
 
---
 
This structure ensures that:
- the schema stays clean and scalable,  
- data redundancy is minimized,  
- and time-varying values (like free disk space) are properly tracked without overwriting historical data.

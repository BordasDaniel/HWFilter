import datetime
from openpyxl import Workbook
import os
import tkinter as tk
from tkinter import messagebox, ttk


# Simple in-memory lookup table to assign stable integer IDs for values
class LookupTable:
    def __init__(self):
        self._map = {}  # normalized value -> id
        self._next = 1

    def normalize(self, v: str) -> str:
        return v.strip().lower()

    def get_or_create(self, v):
        if v is None:
            return None
        # represent dates/times consistently
        if isinstance(v, datetime.date) and not isinstance(v, datetime.datetime):
            key = v.isoformat()
        elif isinstance(v, datetime.datetime):
            key = v.isoformat()
        else:
            key = str(v).strip()
        if key == "":
            return None
        nkey = self.normalize(key)
        if nkey in self._map:
            return self._map[nkey]
        nid = self._next
        self._next += 1
        self._map[nkey] = nid
        return nid


# Global lookup tables for reuse across Filter instances
BRAND_LOOKUP = LookupTable()
MODEL_LOOKUP = LookupTable()
OS_LOOKUP = LookupTable()
USER_LOOKUP = LookupTable()
PCNAME_LOOKUP = LookupTable()
CPU_MODEL_LOOKUP = LookupTable()
CPU_CODE_LOOKUP = LookupTable()
CPU_CODE_FOR_MODEL_LOOKUP = LookupTable()  # Separate lookup for ProcessorModel table
DEVICE_LOOKUP = LookupTable()
FREE_TOTAL_LOOKUP = LookupTable()
DATE_LOOKUP = LookupTable()
TIME_LOOKUP = LookupTable()
NOTES_LOOKUP = LookupTable()


# Relational-style tables (in-memory) mapping id -> record dict
USER_TABLE = {}
PC_TABLE = {}
BRAND_TABLE = {}
MODEL_TABLE = {}
OS_TABLE = {}
DEVICE_TABLE = {}
PROCESSOR_MODEL_TABLE = {}
PROCESSOR_TABLE = {}
LOGIN_TABLE = {}
_LOGIN_NEXT = 1


#region Filter Class

class Filter:

    # Class-level storage for filtered data
    objectsArray: list['Filter'] = []

  


    def __init__(self, login_date, login_time, device_type, pc_name, user, brand,
                 model, installed_ram, cpu_model, cpu_code,
                 operating_system, installation_date, disk,
                 free_total_disk_space, notes):
        self.login_date = login_date
        self.login_time = login_time
        self.device_type = device_type
        self.pc_name = pc_name
        self.user = user
        self.brand = brand
        self.model = model
        self.installed_ram = installed_ram
        self.cpu_code = cpu_code
        self.cpu_model = cpu_model
        self.operating_system = operating_system
        self.installation_date = installation_date
        self.disk = disk
        self.free_total_disk_space = free_total_disk_space
        self.notes = notes
        
        # Create or reuse ids for several fields (the setters already normalize types)
        try:
            self.login_date_id = DATE_LOOKUP.get_or_create(self.login_date)
        except Exception:
            self.login_date_id = None
        try:
            self.login_time_id = TIME_LOOKUP.get_or_create(self.login_time)
        except Exception:
            self.login_time_id = None
        self.user_id = USER_LOOKUP.get_or_create(self.user)
        self.pc_name_id = PCNAME_LOOKUP.get_or_create(self.pc_name)
        self.brand_id = BRAND_LOOKUP.get_or_create(self.brand)
        self.model_id = MODEL_LOOKUP.get_or_create(self.model)
        self.os_id = OS_LOOKUP.get_or_create(self.operating_system)
        self.cpu_model_id = CPU_MODEL_LOOKUP.get_or_create(self.cpu_model)
        self.cpu_code_id = CPU_CODE_LOOKUP.get_or_create(self.cpu_code)
        self.cpu_code_for_model_id = CPU_CODE_FOR_MODEL_LOOKUP.get_or_create(self.cpu_code)  # Separate ID for ProcessorModel
        self.device_type_id = DEVICE_LOOKUP.get_or_create(self.device_type)
        self.free_total_id = FREE_TOTAL_LOOKUP.get_or_create(self.free_total_disk_space)
        self.notes_id = NOTES_LOOKUP.get_or_create(self.notes)

        # Register this object, keeping only the latest per (user_id, pc_name_id)
        Filter.register(self)


    #region Properties

    @property
    def login_date(self):
        return self._login_date
    
    @login_date.setter
    def login_date(self, value):
        parsed = Filter._parse_date(value)
        self._login_date = parsed if parsed is not None else value

    @property
    def login_time(self):
        return self._login_time
    
    @login_time.setter
    def login_time(self, value):
        parsed = Filter._parse_time(value)
        self._login_time = parsed if parsed is not None else value

    @property
    def device_type(self):
        return self._device_type
    
    @device_type.setter
    def device_type(self, value):
        self._device_type = value

    @property
    def pc_name(self):
        return self._pc_name
    
    @pc_name.setter
    def pc_name(self, value):
        self._pc_name = value
    
    @property
    def user(self):
        return self._user
    
    @user.setter
    def user(self, value):
        self._user = value
    
    @property
    def brand(self):
        return self._brand

    @brand.setter
    def brand(self, value):
        self._brand = value

    @property
    def model(self):
        return self._model
    
    @model.setter
    def model(self, value):
        self._model = value

    @property
    def installed_ram(self):
        return self._installed_ram
    
    @installed_ram.setter
    def installed_ram(self, value):
        parsed = Filter._parse_ram(value)
        self._installed_ram = parsed if parsed is not None else value

    @property
    def cpu_model(self):
        return self._cpu_model
    
    @cpu_model.setter
    def cpu_model(self, value):
        self._cpu_model = value

    @property
    def cpu_code(self):
        return self._cpu_code
    
    @cpu_code.setter
    def cpu_code(self, value):
        self._cpu_code = value

    @property
    def operating_system(self):
        return self._operating_system
    
    @operating_system.setter
    def operating_system(self, value):
        self._operating_system = value

    @property
    def installation_date(self):
        return self._installation_date
    
    @installation_date.setter
    def installation_date(self, value):
        parsed = Filter._parse_date(value)
        self._installation_date = parsed if parsed is not None else value
    
    @property
    def disk(self):
        return self._disk
    
    @disk.setter
    def disk(self, value):
        self._disk = value

    @property
    def free_total_disk_space(self):
        return self._free_total_disk_space

    @free_total_disk_space.setter
    def free_total_disk_space(self, value):
        self._free_total_disk_space = value

    @property
    def notes(self):
        return self._notes
    
    @notes.setter
    def notes(self, value):
        self._notes = value


    #endregion

    #region Methods

    # Filter and keep only the latest entry per (user, pc_name) by login_date
    @classmethod
    def register(cls, instance: 'Filter'):
        """Register a Filter instance. If there are existing entries with the same
        (user, pc_name) pair, keep only the one with the latest login_date.

        Matching is case-insensitive on user and pc_name.
        """
        # Prefer numeric ids if available (created during __init__)
        user_id = getattr(instance, 'user_id', None)
        pc_id = getattr(instance, 'pc_name_id', None)
        if user_id is not None and pc_id is not None:
            same = [o for o in cls.objectsArray if getattr(o, 'user_id', None) == user_id and getattr(o, 'pc_name_id', None) == pc_id]
        else:
            # Fallback to case-insensitive string match
            key = (getattr(instance, 'user', '').strip().lower(), getattr(instance, 'pc_name', '').strip().lower())
            same = [o for o in cls.objectsArray if (getattr(o, 'user', '').strip().lower(), getattr(o, 'pc_name', '').strip().lower()) == key]
        if not same:
            cls.objectsArray.append(instance)
            return

        # There are existing objects: include the new instance and choose the latest
        candidates = same + [instance]
        try:
            latest = max(candidates, key=lambda x: x.login_date)
        except Exception:
            # If login_date isn't comparable, fall back to keeping the new instance
            latest = instance

        # Remove all existing same-key objects
        for o in same:
            try:
                cls.objectsArray.remove(o)
            except ValueError:
                pass

        # Ensure latest is in the array
        if latest is instance:
            cls.objectsArray.append(instance)
        else:
            cls.objectsArray.append(latest)

    @classmethod
    def add_object(cls, obj: 'Filter'):
        cls.objectsArray.append(obj)

    @staticmethod
    def _parse_date(value):
        """Parse various date formats into datetime.date. Return None if cannot parse."""
        if value is None:
            return None
        if isinstance(value, datetime.date) and not isinstance(value, datetime.datetime):
            return value
        if isinstance(value, datetime.datetime):
            return value.date()
        if isinstance(value, str):
            s = value.strip()
            if not s:
                return None
            # Try common formats
            fmts = ["%Y-%m-%d", "%Y.%m.%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"]
            for f in fmts:
                try:
                    return datetime.datetime.strptime(s, f).date()
                except Exception:
                    continue
            # try fromisoformat
            try:
                return datetime.date.fromisoformat(s)
            except Exception:
                return None
        return None

    @staticmethod
    def _parse_time(value):
        """Parse time strings into datetime.time. Return None if cannot parse."""
        if value is None:
            return None
        if isinstance(value, datetime.time):
            return value
        if isinstance(value, datetime.datetime):
            return value.time()
        if isinstance(value, str):
            s = value.strip()
            if not s:
                return None
            parts = s.split(":")
            try:
                parts = [int(p) for p in parts]
                if len(parts) == 2:
                    h, m = parts
                    return datetime.time(h, m)
                elif len(parts) >= 3:
                    h, m, sec = parts[:3]
                    return datetime.time(h, m, sec)
            except Exception:
                return None
        return None

    @staticmethod
    def _parse_ram(value):
        """Try to parse installed RAM into an integer number of GB.
        Accepts strings like '8GB', '16 GB', '16384MB', '8', etc.
        Returns int (GB) or None if parsing fails.
        """
        if value is None:
            return None
        if isinstance(value, (int, float)):
            try:
                return int(value)
            except Exception:
                return None
        s = str(value).strip().lower()
        if not s:
            return None
        # Remove common suffixes
        try:
            if s.endswith('gb'):
                return int(float(s[:-2].strip()))
            if s.endswith('g'):
                return int(float(s[:-1].strip()))
            if s.endswith('mb'):
                mb = float(s[:-2].strip())
                return int(mb / 1024)
            if s.endswith('k') or s.endswith('kb'):
                # treat as KB -> GB
                num = float(s.rstrip('kbk').strip())
                return int(num / (1024*1024))
            # plain number
            return int(float(s))
        except Exception:
            return None

    #endregion

#endregion


#region Data Handling Functions
def DataReader():
    file_path = "hw.txt"

    try:
        # A simple progress window
        # First count lines to set the maximum
        try:
            with open(file_path, 'r', encoding='utf-8') as _f:
                total_lines = sum(1 for _ in _f)
        except FileNotFoundError:
            raise

        progress_root = tk.Tk()
        progress_root.title('Loading data')
        progress_root.geometry('400x90')
        progress_root.resizable(False, False)
        progress_root.attributes('-topmost', True)
        tk.Label(progress_root, text=f'Reading data from {file_path}...').pack(pady=(8, 0))
        pb = ttk.Progressbar(progress_root, orient='horizontal', length=360, mode='determinate')
        pb.pack(pady=(8, 8))
        pb['maximum'] = max(1, total_lines)
        progress_root.update()

        # Now process with progress updates
        with open(file_path, "r", encoding="utf-8") as f:
            for idx, line in enumerate(f, start=1):
                line = line.strip()
                if not line:
                    pb['value'] = idx
                    progress_root.update_idletasks()
                    continue
                fields = line.split(";")
                try:
                    # First field expected format YYYY.MM.DD
                    login_date_str = fields[0]
                    y, m, d = map(int, login_date_str.split("."))
                    login_date = datetime.date(y, m, d)
                except (ValueError, IndexError):
                    pb['value'] = idx
                    progress_root.update_idletasks()
                    continue

                # Variables for Filter constructor (guard against short lines)
                # Use empty string if missing
                def safe(i):
                    return fields[i] if i < len(fields) else ""

                login_time = safe(1)
                device_type = safe(2)
                pc_name = safe(3)
                user = safe(4)
                brand = safe(5)
                model = safe(6)
                installed_ram = safe(7)
                cpu_model = safe(8)
                cpu_code = safe(9)
                operating_system = safe(10)
                installation_date = safe(11)
                disk = safe(12)
                free_total_disk_space = safe(13)
                notes = safe(14)

                # Create Filter instance (which auto-registers itself)
                obj = Filter(
                    login_date, login_time, device_type, pc_name, user, brand,
                    model, installed_ram, cpu_model, cpu_code,
                    operating_system, installation_date, disk,
                    free_total_disk_space, notes
                )

                # Populate relational-style tables using ids from the Filter instance
                # Users
                if getattr(obj, 'user_id', None) is not None and obj.user_id not in USER_TABLE:
                    USER_TABLE[obj.user_id] = {'id': obj.user_id, 'name': obj.user}

                # Brand
                if getattr(obj, 'brand_id', None) is not None and obj.brand_id not in BRAND_TABLE:
                    BRAND_TABLE[obj.brand_id] = {'id': obj.brand_id, 'name': obj.brand}

                # Model (attach brand if available)
                if getattr(obj, 'model_id', None) is not None and obj.model_id not in MODEL_TABLE:
                    MODEL_TABLE[obj.model_id] = {'id': obj.model_id, 'brand_id': getattr(obj, 'brand_id', None), 'name': obj.model}

                # OS
                if getattr(obj, 'os_id', None) is not None and obj.os_id not in OS_TABLE:
                    OS_TABLE[obj.os_id] = {'id': obj.os_id, 'name': obj.operating_system}

                # Device type
                if getattr(obj, 'device_type_id', None) is not None and obj.device_type_id not in DEVICE_TABLE:
                    DEVICE_TABLE[obj.device_type_id] = {'id': obj.device_type_id, 'type': obj.device_type}

                # Processor model
                if getattr(obj, 'cpu_code_for_model_id', None) is not None and obj.cpu_code_for_model_id not in PROCESSOR_MODEL_TABLE:
                    PROCESSOR_MODEL_TABLE[obj.cpu_code_for_model_id] = {'id': obj.cpu_code_for_model_id, 'name': obj.cpu_code}

                # Processor - unique processor record for each machine
                # Generate unique processor ID for each machine
                proc_key = f"{obj.pc_name or ''}|{obj.cpu_model or ''}|{obj.cpu_code or ''}"
                proc_id = CPU_CODE_LOOKUP.get_or_create(proc_key)
                
                # Every processor is unique, so always add it
                PROCESSOR_TABLE[proc_id] = {
                    'id': proc_id, 
                    'code': obj.cpu_model,  # ProcessorCode field
                    'model_id': getattr(obj, 'cpu_code_for_model_id', None)  # ProcessorModelID reference
                }

                # PC (reuse by pc_name id)
                if getattr(obj, 'pc_name_id', None) is not None:
                    if obj.pc_name_id not in PC_TABLE:
                        PC_TABLE[obj.pc_name_id] = {
                            'id': obj.pc_name_id,
                            'name': obj.pc_name,
                            'device_id': getattr(obj, 'device_type_id', None),
                            'model_id': getattr(obj, 'model_id', None),
                            'ram_gb': obj.installed_ram,
                            'processor_id': proc_id,
                            'os_id': getattr(obj, 'os_id', None),
                            'os_installation_date': getattr(obj, 'installation_date', None),
                            'disk': obj.disk,  # Disk name directly
                            'note': obj.notes
                        }

                # Login table: create a login row linking user and pc
                global _LOGIN_NEXT
                lid = _LOGIN_NEXT
                LOGIN_TABLE[lid] = {
                    'id': lid,
                    'date': obj.login_date,
                    'time': obj.login_time,
                    'pc_id': getattr(obj, 'pc_name_id', None),
                    'user_id': getattr(obj, 'user_id', None),
                    'free_disk_space': obj.free_total_disk_space  # FreeDiskSpace moved from PC table
                }
                _LOGIN_NEXT += 1

                # update progress
                pb['value'] = idx
                progress_root.update_idletasks()

        # finished
        progress_root.destroy()
    except FileNotFoundError:
        messagebox.showerror("Error", f"Input file not found: {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while reading the file:\n{e}")


def DataFilter(d) -> list[Filter]:
    filtered = [f for f in Filter.objectsArray if f.login_date == d]
    return filtered

def Extractor(d):
    filtered_by_date = DataFilter(d)
    # TODO: Implement the extraction to Excel tables logic here
    try:
        # Build a set of login ids that match the filter
        # d may be a string (YYYY-MM or YYYY-MM-DD) or a date
        # We'll select LOGIN_TABLE rows that match
        import re

        def parse_filter_date(x):
            if x is None:
                return None
            if isinstance(x, datetime.date):
                return x
            if isinstance(x, str):
                x = x.strip()
                # year-month
                if re.match(r"^\d{4}-\d{2}$", x):
                    return x
                try:
                    return Filter._parse_date(x)
                except Exception:
                    return x
            return None

        # Determine selection predicate
        sel = []
        if d is None or str(d).strip() == "":
            sel = list(LOGIN_TABLE.values())
        else:
            dval = parse_filter_date(d)
            if isinstance(dval, str):
                # year-month filter
                sel = [r for r in LOGIN_TABLE.values() if (r.get('date') is not None and getattr(r['date'], 'isoformat', lambda: str(r['date']))()[:7] == dval)]
            elif isinstance(dval, datetime.date):
                sel = [r for r in LOGIN_TABLE.values() if r.get('date') == dval]
            else:
                # fallback: match string prefix
                sval = str(d)
                sel = [r for r in LOGIN_TABLE.values() if r.get('date') and str(r['date']).startswith(sval)]

        # Collect referenced ids
        user_ids = set(r.get('user_id') for r in sel if r.get('user_id') is not None)
        pc_ids = set(r.get('pc_id') for r in sel if r.get('pc_id') is not None)

        # Collect additional referenced ids from PCs
        brand_ids = set()
        model_ids = set()
        os_ids = set()
        device_ids = set()
        processor_ids = set()
        processor_model_ids = set()

        for pid in pc_ids:
            pc = PC_TABLE.get(pid)
            if not pc:
                continue
            if pc.get('brand_id'):
                brand_ids.add(pc.get('brand_id'))
            if pc.get('model_id'):
                model_ids.add(pc.get('model_id'))
            if pc.get('os_id'):
                os_ids.add(pc.get('os_id'))
            if pc.get('device_id'):
                device_ids.add(pc.get('device_id'))
            if pc.get('processor_id'):
                processor_ids.add(pc.get('processor_id'))
            if pc.get('processor_id'):
                # processor -> model
                proc = PROCESSOR_TABLE.get(pc.get('processor_id'))
                if proc and proc.get('model_id'):
                    processor_model_ids.add(proc.get('model_id'))

        # Also include models' brands
        for mid in list(model_ids):
            m = MODEL_TABLE.get(mid)
            if m and m.get('brand_id'):
                brand_ids.add(m.get('brand_id'))

        # Create workbook and sheets
        wb = Workbook()
        # remove default
        default = wb.active
        wb.remove(default)

        def write_table(name, headers, rows):
            ws = wb.create_sheet(title=name[:31])
            ws.append(headers)
            for row in rows:
                ws.append(row)

        # Login sheet
        login_rows = []
        for r in sel:
            login_rows.append([r.get('id'), r.get('date'), r.get('time'), r.get('pc_id'), r.get('user_id'), r.get('free_disk_space')])
        write_table('Login', ['ID', 'Date', 'Time', 'PC_ID', 'User_ID', 'FreeDiskSpace'], login_rows)

        # User sheet
        user_rows = []
        for uid in sorted(user_ids):
            u = USER_TABLE.get(uid)
            if u:
                user_rows.append([u.get('id'), u.get('name')])
        write_table('User', ['ID', 'Name'], user_rows)

        # PC sheet
        pc_rows = []
        for pid in sorted(pc_ids):
            p = PC_TABLE.get(pid)
            if not p:
                continue
            pc_rows.append([
                p.get('id'), p.get('name'), p.get('device_id'), p.get('model_id'), p.get('ram_gb'),
                p.get('processor_id'), p.get('os_id'), p.get('os_installation_date'), 
                p.get('disk'), p.get('note')
            ])
        write_table('Pc', ['ID', 'Name', 'DeviceID', 'ModelID', 'RAM', 'ProcessorID', 'OperationSystemID', 'OperationSystemInstallationDate', 'Disk', 'Note'], pc_rows)

        # Device
        device_rows = [[v.get('id'), v.get('type')] for k, v in DEVICE_TABLE.items() if k in device_ids]
        write_table('Device', ['ID', 'Type'], device_rows)

        # Model
        model_rows = []
        for mid in sorted(model_ids):
            m = MODEL_TABLE.get(mid)
            if m:
                model_rows.append([m.get('id'), m.get('brand_id'), m.get('name')])
        write_table('Model', ['ID', 'BrandID', 'Name'], model_rows)

        # Brand
        brand_rows = [[b.get('id'), b.get('name')] for k, b in BRAND_TABLE.items() if k in brand_ids]
        write_table('Brand', ['ID', 'Name'], brand_rows)

        # OperationSystem
        os_rows = [[o.get('id'), o.get('name')] for k, o in OS_TABLE.items() if k in os_ids]
        write_table('OperationSystem', ['ID', 'Name'], os_rows)

        # ProcessorModel
        pm_rows = [[m.get('id'), m.get('name')] for k, m in PROCESSOR_MODEL_TABLE.items() if k in processor_model_ids]
        write_table('ProcessorModel', ['ID', 'Name'], pm_rows)

        # Processor
        proc_rows = [[p.get('id'), p.get('code'), p.get('model_id')] for k, p in PROCESSOR_TABLE.items() if k in processor_ids]
        write_table('Processor', ['ID', 'ProcessorCode', 'ProcessorModelID'], proc_rows)

        # Save workbook
        dstr = ''
        if d is None or str(d).strip() == '':
            dstr = 'all'
        else:
            try:
                if isinstance(d, str) and re.match(r"^\d{4}-\d{2}$", d):
                    dstr = d
                else:
                    dp = Filter._parse_date(d)
                    dstr = dp.isoformat() if dp else str(d)
            except Exception:
                dstr = str(d)

        output_file = os.path.join(os.getcwd(), f"hw_relational_{dstr}.xlsx")
        wb.save(output_file)
        messagebox.showinfo('Success', f'Excel exported: {output_file}')
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

#endregion

#region Tkinter GUI Functions

def TkinterMain():
    # --- Enable mouse wheel scrolling for the ticket list ---
    def _on_mousewheel(event):
        try:
            if hasattr(event, 'delta') and event.delta:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            elif hasattr(event, 'num') and event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif hasattr(event, 'num') and event.num == 5:
                canvas.yview_scroll(1, "units")
        except tk.TclError:
            return

    def _bind_canvas_wheel(_ev=None):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", _on_mousewheel)
        canvas.bind_all("<Button-5>", _on_mousewheel)

    def _unbind_canvas_wheel(_ev=None):
        try:
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")
        except Exception:
            pass

    # canvas is defined below, so bind after its creation
    root = tk.Tk()
    root.title("HW Filter")
    # Make the window narrower as requested
    root.geometry("700x600")
    root.configure(bg="#eeeeee")

    # --- Search bar at the top ---
    search_frame = tk.Frame(root, bg="#dddddd")
    search_frame.pack(fill="x", padx=10, pady=(12, 4))
    tk.Label(search_frame, text="Search", bg="#dddddd", font=("Arial", 11)).pack(side="left", padx=(8, 4))
    search_var = tk.StringVar()
    search_entry = tk.Entry(search_frame, textvariable=search_var, font=("Arial", 11), width=40)
    search_entry.pack(side="left", padx=4)

    # --- Ticket list (scrollable) ---
    list_outer = tk.Frame(root, bg="#cccccc")
    list_outer.pack(fill="both", expand=True, padx=10, pady=(0, 0))
    canvas = tk.Canvas(list_outer, bg="#cccccc", highlightthickness=0, height=260)
    scrollbar = tk.Scrollbar(list_outer, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg="#cccccc")
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Bind mouse wheel events after canvas is defined
    canvas.bind('<Enter>', _bind_canvas_wheel)
    canvas.bind('<Leave>', _unbind_canvas_wheel)



    # --- Prepare ticket data ---
    filtered = Filter.objectsArray

    # Prepare rows for display: (login_date, user, pc_name)
    filter_rows = [(t.login_date.isoformat(), t.user, t.pc_name) for t in filtered]

    item_frames = []
    selected_index = [None]
    PAGE_SIZE = 50
    current_page = [0]  # mutable for closure
    filtered_rows = list(filter_rows)  # always up-to-date filtered list

    def on_select(idx):
        for i, fr in enumerate(item_frames):
            fr.config(bg="#ffffff")
            for w in fr.winfo_children():
                w.config(bg="#ffffff", fg="#000000")
        item_frames[idx].config(bg="#3399ff")
        for w in item_frames[idx].winfo_children():
            w.config(bg="#3399ff", fg="#ffffff")
        selected_index[0] = idx
        # idx 0 is header, so only for idx > 0
        if idx > 0:
            # Find the correct row in the current page
            page = current_page[0]
            start = page * PAGE_SIZE
            # Use filtered_rows for current search/page
            if start + (idx-1) < len(filtered_rows):
                date_val = filtered_rows[start + (idx-1)][0]
                # Extract only year and month (YYYY-MM) from the date
                year_month = date_val[:7] if date_val and len(date_val) >= 7 else date_val
                try:
                    dateselector.delete(0, tk.END)
                    dateselector.insert(0, year_month)
                except Exception:
                    pass

    def fill_listbox(rows, page=0):
        for fr in item_frames:
            fr.destroy()
        item_frames.clear()
        selected_index[0] = None
        # Header row: three columns (Login Date, User, PC Name) with reduced widths
        header = tk.Frame(scrollable_frame, bg="#dddddd", bd=2, relief="flat")
        tk.Label(header, text="Login Date", font=("Arial", 11, "bold"), bg="#dddddd", anchor="w", width=12).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        tk.Label(header, text="User", font=("Arial", 11, "bold"), bg="#dddddd", anchor="w", width=20).grid(row=0, column=1, sticky="w", padx=8, pady=4)
        tk.Label(header, text="PC Name", font=("Arial", 11, "bold"), bg="#dddddd", anchor="w", width=20).grid(row=0, column=2, sticky="w", padx=8, pady=4)
        header.pack(fill="x", padx=4, pady=(2,2))
        item_frames.append(header)
        # Data rows (paginated)
        start = page * PAGE_SIZE
        end = start + PAGE_SIZE
        page_rows = rows[start:end]
        for idx, (login_date, user, pc_name) in enumerate(page_rows):
            fr = tk.Frame(scrollable_frame, bg="#ffffff", bd=2, relief="groove")
            lbl_date = tk.Label(fr, text=login_date, font=("Arial", 11), bg="#ffffff", anchor="w", width=12)
            lbl_date.grid(row=0, column=0, sticky="w", padx=8, pady=4)
            lbl_user = tk.Label(fr, text=user, font=("Arial", 11), bg="#ffffff", anchor="w", width=20, justify="left")
            lbl_user.grid(row=0, column=1, sticky="w", padx=8, pady=4)
            lbl_pc = tk.Label(fr, text=pc_name, font=("Arial", 11), bg="#ffffff", anchor="w", width=20)
            lbl_pc.grid(row=0, column=2, sticky="w", padx=8, pady=4)
            fr.pack(fill="x", padx=4, pady=2)
            fr.bind("<Button-1>", lambda e, i=idx+1: on_select(i))
            lbl_date.bind("<Button-1>", lambda e, i=idx+1: on_select(i))
            lbl_user.bind("<Button-1>", lambda e, i=idx+1: on_select(i))
            lbl_pc.bind("<Button-1>", lambda e, i=idx+1: on_select(i))
            item_frames.append(fr)
        # Pagination controls
        total_pages = max(1, (len(rows) + PAGE_SIZE - 1) // PAGE_SIZE)
        pag_frame = tk.Frame(scrollable_frame, bg="#eeeeee")
        pag_frame.pack(fill="x", pady=(6, 2))
        prev_btn = tk.Button(pag_frame, text="< Previous", state=("normal" if page > 0 else "disabled"), command=lambda p=page: goto_page(p-1))
        prev_btn.pack(side="left", padx=8)
        page_label = tk.Label(pag_frame, text=f"Page {page+1} / {total_pages}", font=("Arial", 10), bg="#eeeeee")
        page_label.pack(side="left", padx=8)
        next_btn = tk.Button(pag_frame, text="Next >", state=("normal" if page < total_pages-1 else "disabled"), command=lambda p=page: goto_page(p+1))
        next_btn.pack(side="left", padx=8)
        item_frames.append(pag_frame)

    def goto_page(page):
        current_page[0] = page
        fill_listbox(filtered_rows, page)
        # Scroll to top of canvas after page change
        canvas.yview_moveto(0)

    fill_listbox(filter_rows, 0)

    def on_search(*args):
        q = search_var.get().strip()
        nonlocal filtered_rows
        if not q:
            filtered_rows = [row for row in filter_rows]
        else:
            # Match year-month only (YYYY-MM)
            import re
            ym_pattern = r"^\d{4}-\d{2}$"
            if re.match(ym_pattern, q):
                filtered_rows = [row for row in filter_rows if (row[0] and row[0][:7]) == q]
            else:
                # if user types partial, try to match start of YYYY-MM
                filtered_rows = [row for row in filter_rows if row[0] and row[0].startswith(q)]
        current_page[0] = 0
        fill_listbox(filtered_rows, current_page[0])

    search_var.trace_add('write', on_search)

    # --- Controls at the bottom: only two buttons as requested ---
    btn_frame = tk.Frame(root, bg="#eeeeee")
    btn_frame.pack(pady=(8, 16))

    tk.Label(btn_frame, text="Enter date (YYYY-MM):").grid(row=0, column=0, padx=8, pady=4)
    dateselector = tk.Entry(btn_frame)
    dateselector.grid(row=0, column=1, padx=8, pady=4)

    def parse_date_input(s: str):
        s = s.strip()
        if not s:
            return None
        # Accept year-month like '2024-02' and return as string for extractor
        import re
        if re.match(r"^\d{4}-\d{2}$", s):
            return s
        fmts = ['%Y-%m-%d', '%Y/%m/%d', '%d-%m-%Y', '%d/%m/%Y']
        for f in fmts:
            try:
                return datetime.datetime.strptime(s, f).date()
            except Exception:
                continue
        try:
            return datetime.date.fromisoformat(s)
        except Exception:
            return None

    def on_generate():
        s = dateselector.get()
        d = parse_date_input(s)
        # Call Extractor if it exists; otherwise inform the user
        try:
            Extractor(d)
            # Clear the date input field after successful export
            dateselector.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror('Error', str(e))

    button = tk.Button(btn_frame, text="Generate Excel", command=on_generate)
    button.grid(row=1, column=0, padx=8, pady=4)

    close_btn = tk.Button(btn_frame, text="Close", command=root.destroy, width=16, bg="#888888", fg="#fff")
    close_btn.grid(row=1, column=1, padx=8, pady=4)

    root.mainloop()

#endregion

#endregion

def Main():
    DataReader()
    TkinterMain()


if __name__ == "__main__":
    Main()
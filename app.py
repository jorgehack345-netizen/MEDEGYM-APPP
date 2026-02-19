# -*- coding: utf-8 -*-
import os, sys, json, webbrowser
from datetime import date, datetime

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from excel_backend import GymWorkbook


def _resource_path(rel):
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base, rel)

# Load config
CFG = {"gym_name":"Gimnasio","whatsapp_prefix":"+52","message_template":"Hola {nombre}, tu membresía vence el {fecha}. Por favor realiza tu pago para mantener tu acceso. ¡Gracias!"}
try:
    with open(_resource_path('config.json'), 'r', encoding='utf-8') as _cf:
        CFG.update(json.load(_cf))
except Exception:
    pass

APP_TITLE = f"{CFG.get('gym_name','Gimnasio')} - Control de Gimnasio"

class GymApp(ttk.Frame):
    def __init__(self, master, xlsx_path=None):
        super().__init__(master)
        self.master.title(APP_TITLE)
        self.master.geometry("1150x700")
        # icon/logo
        try:
            from tkinter import PhotoImage
            logo_path = _resource_path('logo.png')
            if os.path.exists(logo_path):
                self.master.iconphoto(True, PhotoImage(file=logo_path))
        except Exception:
            pass

        self.pack(fill=tk.BOTH, expand=True)

        self.wb = None
        self.xlsx_path = xlsx_path or os.path.join(os.getcwd(), 'Control_Gimnasio_Pagos.xlsx')
        self.create_widgets()
        self.load_wb(self.xlsx_path)

    def create_widgets(self):
        top = ttk.Frame(self)
        top.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)

        ttk.Label(top, text="Archivo Excel:").pack(side=tk.LEFT)
        self.file_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.file_var, width=80).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Cambiar...", command=self.change_file).pack(side=tk.LEFT)
        ttk.Button(top, text="Refrescar", command=self.refresh_all).pack(side=tk.LEFT, padx=6)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        self.tab_resumen = ttk.Frame(self.nb)
        self.tab_miembros = ttk.Frame(self.nb)
        self.tab_pagos = ttk.Frame(self.nb)
        self.tab_alertas = ttk.Frame(self.nb)

        self.nb.add(self.tab_resumen, text="Resumen")
        self.nb.add(self.tab_miembros, text="Miembros")
        self.nb.add(self.tab_pagos, text="Pagos")
        self.nb.add(self.tab_alertas, text="Alertas")

        # Resumen
        self.lbl_total = ttk.Label(self.tab_resumen, text="Total: 0", font=("Segoe UI", 11, "bold"))
        self.lbl_total.pack(anchor=tk.W, padx=8, pady=6)
        self.lbl_estados = ttk.Label(self.tab_resumen, text="Estados: -", font=("Segoe UI", 10))
        self.lbl_estados.pack(anchor=tk.W, padx=8)

        # Miembros grid
        cols_m = ("ID","Nombre","Teléfono","Plan","Monto","F.Inicio","F.Ult.Pago","Próx.Venc","Días","Estado")
        self.tv_m = ttk.Treeview(self.tab_miembros, columns=cols_m, show='headings', height=18)
        for c in cols_m:
            self.tv_m.heading(c, text=c)
            self.tv_m.column(c, width=110 if c!="Nombre" else 180, anchor=tk.CENTER)
        self.tv_m.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        frm_add = ttk.Labelframe(self.tab_miembros, text="Agregar miembro")
        frm_add.pack(fill=tk.X, padx=8, pady=6)
        self.id_var = tk.StringVar(); self.nom_var = tk.StringVar(); self.tel_var = tk.StringVar()
        self.plan_var = tk.StringVar(value='Mensual'); self.monto_var = tk.StringVar(value='0')
        self.ini_var = tk.StringVar(value=date.today().strftime('%d/%m/%Y'))

        for i,(lbl,var,w) in enumerate([
            ("ID",self.id_var,10),("Nombre",self.nom_var,28),("Teléfono",self.tel_var,16),
            ("Plan",self.plan_var,12),("Monto",self.monto_var,10),("Fecha inicio (dd/mm/aaaa)",self.ini_var,16)
        ]):
            ttk.Label(frm_add, text=lbl+":").grid(row=0, column=2*i, padx=5, pady=4, sticky=tk.E)
            if lbl=="Plan":
                cb = ttk.Combobox(frm_add, textvariable=var, width=w, values=["Semanal","Quincenal","Mensual","Trimestral","Semestral","Anual"])
                cb.grid(row=0, column=2*i+1, padx=5, pady=4)
            else:
                ttk.Entry(frm_add, textvariable=var, width=w).grid(row=0, column=2*i+1, padx=5, pady=4)

        ttk.Button(frm_add, text="Agregar", command=self.add_member).grid(row=0, column=12, padx=8)

        # Pagos grid
        cols_p = ("ID Miembro","Nombre","Fecha pago","Monto","Método","Notas")
        self.tv_p = ttk.Treeview(self.tab_pagos, columns=cols_p, show='headings', height=16)
        for c in cols_p:
            self.tv_p.heading(c, text=c)
            self.tv_p.column(c, width=120 if c!="Notas" else 220, anchor=tk.CENTER)
        self.tv_p.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        frm_pay = ttk.Labelframe(self.tab_pagos, text="Registrar pago")
        frm_pay.pack(fill=tk.X, padx=8, pady=6)
        self.pid_var = tk.StringVar(); self.pname_var = tk.StringVar(); self.pdate_var = tk.StringVar(value=date.today().strftime('%d/%m/%Y'))
        self.pamt_var = tk.StringVar(value='0'); self.pmet_var = tk.StringVar(value='Efectivo'); self.pnote_var = tk.StringVar()
        for i,(lbl,var,w) in enumerate([
            ("ID Miembro",self.pid_var,10),("Nombre",self.pname_var,24),("Fecha pago (dd/mm/aaaa)",self.pdate_var,16),
            ("Monto",self.pamt_var,10),("Método",self.pmet_var,14),("Notas",self.pnote_var,30)
        ]):
            ttk.Label(frm_pay, text=lbl+":").grid(row=0, column=2*i, padx=5, pady=4, sticky=tk.E)
            if lbl=="Método":
                cb = ttk.Combobox(frm_pay, textvariable=var, width=w, values=["Efectivo","Tarjeta","Transferencia"])
                cb.grid(row=0, column=2*i+1, padx=5, pady=4)
            else:
                ttk.Entry(frm_pay, textvariable=var, width=w).grid(row=0, column=2*i+1, padx=5, pady=4)
        ttk.Button(frm_pay, text="Guardar pago", command=self.add_payment).grid(row=0, column=12, padx=8)

        # Alertas grid
        cols_a = ("ID","Nombre","Teléfono","Plan","Monto","F.Inicio","F.Ult.Pago","Próx.Venc","Días","Estado","Mensaje","WhatsApp")
        self.tv_a = ttk.Treeview(self.tab_alertas, columns=cols_a, show='headings', height=16)
        for c,w in zip(cols_a,[80,160,120,100,80,100,110,110,60,100,260,220]):
            self.tv_a.heading(c, text=c)
            self.tv_a.column(c, width=w, anchor=tk.CENTER)
        self.tv_a.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        btns_a = ttk.Frame(self.tab_alertas); btns_a.pack(fill=tk.X, padx=8, pady=6)
        ttk.Button(btns_a, text="Abrir WhatsApp del seleccionado", command=self.open_whatsapp).pack(side=tk.LEFT)

    def load_wb(self, path):
        try:
            self.wb = GymWorkbook(path)
            self.file_var.set(path)
            self.refresh_all()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir/crear el libro.\n{e}")

    def change_file(self):
        p = filedialog.askopenfilename(title="Selecciona Excel", filetypes=[["Excel","*.xlsx"]])
        if not p:
            return
        self.load_wb(p)

    def refresh_all(self):
        if not self.wb: return
        data = self.wb.get_summary()
        self.lbl_total.config(text=f"Total: {data['total']}")
        self.lbl_estados.config(text=(
            f"Vencidos: {data['Vencido']} | Vence hoy: {data['Vence hoy']} | Por vencer: {data['Por vencer']} | Al día: {data['Al día']}"
        ))
        # Miembros
        for i in self.tv_m.get_children(): self.tv_m.delete(i)
        for m in data['miembros']:
            self.tv_m.insert('', 'end', values=(
                m['id'], m['nombre'], m['telefono'], m['plan'], f"${m['monto']:.2f}",
                m['f_inicio'].strftime('%d/%m/%Y') if m['f_inicio'] else '',
                m['f_ult_pago'].strftime('%d/%m/%Y') if m['f_ult_pago'] else '',
                m['prox_venc'].strftime('%d/%m/%Y') if m['prox_venc'] else '',
                m['dias'] if m['dias'] is not None else '', m['estado']
            ))
        # Pagos
        for i in self.tv_p.get_children(): self.tv_p.delete(i)
        for p in data['pagos'][-300:]:
            self.tv_p.insert('', 'end', values=(
                p['id'], p['nombre'], p['fecha'].strftime('%d/%m/%Y'), f"${p['monto']:.2f}", p['metodo'], p['notas']
            ))
        # Alertas
        for i in self.tv_a.get_children(): self.tv_a.delete(i)
        for a in data['alertas']:
            self.tv_a.insert('', 'end', values=(
                a['id'], a['nombre'], a['telefono'], a['plan'], f"${a['monto']:.2f}",
                a['f_inicio'].strftime('%d/%m/%Y') if a['f_inicio'] else '',
                a['f_ult_pago'].strftime('%d/%m/%Y') if a['f_ult_pago'] else '',
                a['prox_venc'].strftime('%d/%m/%Y'), a['dias'], a['estado'], a['mensaje'], a['wa_link']
            ))

    def add_member(self):
        try:
            m = {
                'id': self.id_var.get().strip(),
                'nombre': self.nom_var.get().strip(),
                'telefono': self.tel_var.get().strip(),
                'plan': self.plan_var.get().strip(),
                'monto': float(self.monto_var.get()),
                'f_inicio': datetime.strptime(self.ini_var.get().strip(), '%d/%m/%Y').date(),
            }
        except Exception as e:
            messagebox.showerror("Datos inválidos", f"Revisa los campos.\n{e}")
            return
        try:
            self.wb.add_member(m)
            self.wb.save()
            self.refresh_all()
            messagebox.showinfo("Listo", "Miembro agregado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def add_payment(self):
        try:
            p = {
                'id': self.pid_var.get().strip(),
                'nombre': self.pname_var.get().strip(),
                'fecha': datetime.strptime(self.pdate_var.get().strip(), '%d/%m/%Y').date(),
                'monto': float(self.pamt_var.get()),
                'metodo': self.pmet_var.get().strip(),
                'notas': self.pnote_var.get().strip(),
            }
        except Exception as e:
            messagebox.showerror("Datos inválidos", f"Revisa los campos.\n{e}")
            return
        try:
            self.wb.add_payment(p)
            self.wb.save()
            self.refresh_all()
            messagebox.showinfo("Listo", "Pago registrado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def open_whatsapp(self):
        sel = self.tv_a.selection()
        if not sel:
            messagebox.showinfo("Info", "Selecciona una fila en Alertas.")
            return
        vals = self.tv_a.item(sel[0], 'values')
        link = vals[-1]
        try:
            webbrowser.open(link)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el navegador.\n{e}")


def main():
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use('clam')
    except Exception:
        pass
    app = GymApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()

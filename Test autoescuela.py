import customtkinter as ctk
import pandas as pd
import random
from pathlib import Path
import os
from tkinter import messagebox
import tkinter as tk
from PIL import Image

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class ExamSystemGUI:
    def __init__(self):
        self.estadisticas_file = 'estadisticas.xlsx'
        self.current_question_index = 0
        self.exam_questions = []
        self.exam_answers = []
        self.exam_results = []
        self.setup_main_window()
        self.create_main_interface()

    def setup_main_window(self):
        self.root = ctk.CTk()
        self.root.title("Sistema de Exámenes Profesional")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)

    def create_main_interface(self):
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        self.create_header()
        self.content_frame = ctk.CTkFrame(self.main_frame)
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=(0,10))
        self.show_home_screen()

    def create_header(self):
        hf = ctk.CTkFrame(self.main_frame, height=80)
        hf.pack(fill="x", padx=20, pady=(20,20))
        hf.pack_propagate(False)
        ctk.CTkLabel(hf, text="Sistema de Exámenes Profesional",
            font=ctk.CTkFont(size=26, weight="bold")).pack(side="left", padx=20, pady=20)
        nf = ctk.CTkFrame(hf, fg_color="transparent")
        nf.pack(side="right", padx=20, pady=20)
        for txt, cmd in [("Inicio", self.show_home_screen),
                         ("Estadísticas", self.show_statistics_screen),
                         ("Configuración", self.show_settings_screen)]:
            ctk.CTkButton(nf, text=txt, command=cmd, width=120, height=35,
                font=ctk.CTkFont(size=15)).pack(side="left", padx=5)

    def clear_content(self):
        for w in self.content_frame.winfo_children():
            w.destroy()

    def show_home_screen(self):
        self.clear_content()
        ctk.CTkLabel(self.content_frame, text="Selecciona el tipo de examen",
            font=ctk.CTkFont(size=22, weight="bold")).pack(pady=(10,10))
        of = ctk.CTkFrame(self.content_frame)
        of.pack(fill="both", expand=True, padx=40, pady=10)
        self.create_exam_options(of)
        self.create_info_panel(of)

    def create_exam_options(self, parent):
        lf = ctk.CTkScrollableFrame(parent)
        lf.pack(side="left", fill="both", expand=True, padx=(20,10), pady=20)
        ctk.CTkLabel(lf, text="Tipos de Examen",
            font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(10,15))
        for title, desc, cmd in [
            ("Examen por Tema", "Selecciona un tema concreto", self.start_single_topic_exam),
            ("Examen Multi-Tema", "Combina varios temas", self.start_multi_topic_exam),
            ("Examen Completo", "Todos los temas disponibles", self.start_complete_exam),
            ("Preguntas más Falladas", "Repasa lo que más fallas", self.start_failed_questions_exam)
        ]:
            fr = ctk.CTkFrame(lf)
            fr.pack(fill="x", padx=20, pady=8)
            ctk.CTkButton(fr, text=title, command=cmd, height=40,
                font=ctk.CTkFont(size=14, weight="bold")).pack(fill="x", padx=15, pady=(12,4))
            ctk.CTkLabel(fr, text=desc, font=ctk.CTkFont(size=12),
                text_color="gray60").pack(padx=15, pady=(0,12))

    def create_info_panel(self, parent):
        rf = ctk.CTkFrame(parent, width=300)
        rf.pack(side="right", fill="y", padx=(10,20), pady=20)
        rf.pack_propagate(False)
        ctk.CTkLabel(rf, text="Estado del Sistema",
            font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(20,20))
        self.update_system_info(rf)
        ctk.CTkButton(rf, text="Actualizar",
            command=lambda: self.update_system_info(rf), height=35).pack(side="bottom", padx=20, pady=20)

    def update_system_info(self, parent):
        for w in parent.winfo_children()[1:-1]:
            w.destroy()
        temas = self.get_available_topics()
        inf = ctk.CTkFrame(parent, fg_color="transparent")
        inf.pack(fill="x", padx=20, pady=10)
        total = sum(len(self.load_questions(t)) for t in temas)
        stats = self.load_statistics()
        attempts = sum(s['intentos'] for s in stats.values())
        for txt in [f"Temas: {len(temas)}", f"Preguntas: {total}", f"Intentos: {attempts}"]:
            ctk.CTkLabel(inf, text=txt, font=ctk.CTkFont(size=13)).pack(anchor="w", pady=2)
        if temas:
            ctk.CTkLabel(inf, text="Temas:", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(15,5))
            tf = ctk.CTkScrollableFrame(inf, height=180)
            tf.pack(fill="x")
            for t in sorted(temas)[:15]:
                ctk.CTkLabel(tf, text=f"• {t}", font=ctk.CTkFont(size=12),
                    text_color="gray60").pack(anchor="w", padx=10, pady=2)

    def load_questions_from_file(self, archivo):
        preguntas = []
        tema_nombre = archivo.stem
        try:
            df = pd.read_excel(archivo) if archivo.suffix == '.xlsx' else pd.read_csv(archivo, sep=';')
            df.columns = [c.strip() for c in df.columns]
            col_map = {}
            for col in df.columns:
                cn = col.lower().replace('ó','o').replace('í','i').replace('é','e').replace('á','a')
                if 'pregunta' in cn: col_map['pregunta'] = col
                elif cn in ['correcta','respuesta']: col_map['correcta'] = col
                elif cn in ['imagen','image','img','foto']: col_map['imagen'] = col
                elif 'opcion' in cn or 'option' in cn:
                    if cn.endswith('a'): col_map['opcionA'] = col
                    elif cn.endswith('b'): col_map['opcionB'] = col
                    elif cn.endswith('c'): col_map['opcionC'] = col
            if 'pregunta' not in col_map: return []
            opcols = [col_map.get(f'opcion{l}') for l in ['A','B','C']]
            pcol = col_map['pregunta']
            ccol = col_map.get('correcta')
            icol = col_map.get('imagen')
            for _, row in df.iterrows():
                try:
                    if pd.isna(row[pcol]) or str(row[pcol]).strip() == '': continue
                    opts = []
                    for i, c in enumerate(opcols):
                        opts.append(str(row[c]).strip() if c and not pd.isna(row[c]) else f'Opción {chr(65+i)}')
                    while len(opts) < 3: opts.append(f'Opción {chr(65+len(opts))}')
                    correcta = 'A'
                    if ccol and not pd.isna(row[ccol]):
                        cr = str(row[ccol]).upper().strip()
                        if cr in ['A','B','C']: correcta = cr
                        elif cr in ['1','2','3']: correcta = chr(64+int(cr))
                    img_ruta = None
                    if icol and icol in df.columns and not pd.isna(row[icol]):
                        ruta = str(row[icol]).strip()
                        for cand in [Path(ruta), Path('imagenes')/ruta, Path('imagenes')/Path(ruta).name]:
                            if cand.exists(): img_ruta = str(cand); break
                    preguntas.append({'tema': tema_nombre, 'pregunta': str(row[pcol]).strip(),
                        'opciones': opts[:3], 'correcta': correcta, 'imagen': img_ruta})
                except: continue
        except Exception as e:
            print(f"Error {tema_nombre}: {e}")
        return preguntas

    def load_questions(self, tema=None):
        dp = Path('datos')
        if not dp.exists(): return []
        todas = []
        if tema:
            for ext in ['.xlsx','.csv']:
                a = dp/f"{tema}{ext}"
                if a.exists(): todas.extend(self.load_questions_from_file(a)); break
        else:
            for a in list(dp.glob('*.xlsx'))+list(dp.glob('*.csv')):
                todas.extend(self.load_questions_from_file(a))
        return todas

    def get_available_topics(self):
        dp = Path('datos')
        if not dp.exists(): return []
        return sorted([a.stem for a in list(dp.glob('*.xlsx'))+list(dp.glob('*.csv'))])

    def show_question_screen(self):
        if self.current_question_index >= len(self.exam_questions):
            self.show_exam_results(); return
        self.clear_content()
        q = self.exam_questions[self.current_question_index]
        qn = self.current_question_index + 1
        tq = len(self.exam_questions)
        tb = ctk.CTkFrame(self.content_frame, height=60)
        tb.pack(fill="x", padx=20, pady=(10,10))
        if qn > 1:
            ctk.CTkButton(tb, text="Anterior", command=self.previous_question,
                width=100, height=38).pack(side="left", padx=5)
        ctk.CTkLabel(tb, text=f"Pregunta {qn} de {tq}",
            font=ctk.CTkFont(size=15, weight="bold")).pack(side="left", padx=20, expand=True)
        self.result_label = ctk.CTkLabel(tb, text="", font=ctk.CTkFont(size=15, weight="bold"), width=300)
        self.result_label.pack(side="left", padx=10)
        rb = ctk.CTkFrame(tb, fg_color="transparent")
        rb.pack(side="right", padx=5)
        ctk.CTkButton(rb, text="Salir", command=self.confirm_exit_exam,
            width=90, height=38, fg_color="#BF616A").pack(side="right", padx=5)
        if qn < tq:
            ctk.CTkButton(rb, text="Siguiente", command=self.next_question,
                width=110, height=38, fg_color="#4ECDC4").pack(side="right", padx=5)
        else:
            ctk.CTkButton(rb, text="Finalizar", command=self.finish_exam,
                width=110, height=38, fg_color="#1DD1A1",
                font=ctk.CTkFont(weight="bold")).pack(side="right", padx=5)
        ms = ctk.CTkScrollableFrame(self.content_frame)
        ms.pack(fill="both", expand=True, padx=20, pady=(0,20))
        qfr = ctk.CTkFrame(ms)
        qfr.pack(fill="x", padx=10, pady=10)
        qtb = ctk.CTkTextbox(qfr, height=100, font=ctk.CTkFont(size=20), wrap="word")
        qtb.pack(fill="x", padx=15, pady=15)
        qtb.insert("1.0", f"Pregunta: {q['pregunta']}")
        qtb.configure(state="disabled")
        if q.get('imagen'):
            try:
                img = Image.open(q['imagen'])
                img.thumbnail((400, 300), Image.LANCZOS)
                w, h = img.size
                ci = ctk.CTkImage(light_image=img, dark_image=img, size=(w,h))
                lbl = ctk.CTkLabel(ms, image=ci, text="")
                lbl.image = ci
                lbl.pack(pady=(5,10))
            except:
                ctk.CTkLabel(ms, text=f"No se pudo cargar: {q['imagen']}",
                    font=ctk.CTkFont(size=11), text_color="gray60").pack(pady=5)
        self.answer_var = ctk.StringVar(value="")
        opts = q.get('opciones', ['','',''])
        while len(opts) < 3: opts.append(f"Opción {chr(65+len(opts))}")
        for i in range(3):
            letra = chr(65+i)
            texto = str(opts[i]).strip()
            if not texto or texto == "nan": texto = f"Opción {letra}"
            ofr = ctk.CTkFrame(ms, corner_radius=10)
            ofr.pack(fill="x", padx=10, pady=8)
            ctk.CTkRadioButton(ofr, text=letra, variable=self.answer_var, value=letra,
                font=ctk.CTkFont(size=20, weight="bold"), height=45, width=50,
                command=self.check_answer_immediately).pack(side="left", padx=(15,10), pady=12)
            ctk.CTkLabel(ofr, text=texto, font=ctk.CTkFont(size=17),
                wraplength=900, justify="left", anchor="w").pack(
                side="left", fill="x", expand=True, padx=(0,15), pady=12)
        self.restore_answer()

    def check_answer_immediately(self):
        if not hasattr(self,'result_label'): return
        sel = self.answer_var.get()
        if not sel: self.result_label.configure(text=""); return
        cor = self.exam_questions[self.current_question_index]['correcta']
        if sel == cor:
            self.result_label.configure(text="CORRECTO!", text_color="#A3BE8C")
        else:
            self.result_label.configure(text=f"Incorrecto. Correcta: {cor}", text_color="#BF616A")

    def previous_question(self):
        if self.current_question_index > 0:
            self.save_current_answer()
            self.current_question_index -= 1
            self.show_question_screen()
            self.restore_answer()
            self.check_answer_immediately()

    def next_question(self):
        self.save_current_answer()
        self.current_question_index += 1
        self.show_question_screen()
        self.restore_answer()
        self.check_answer_immediately()

    def save_current_answer(self):
        while len(self.exam_answers) <= self.current_question_index:
            self.exam_answers.append("")
        self.exam_answers[self.current_question_index] = self.answer_var.get()

    def restore_answer(self):
        if self.current_question_index < len(self.exam_answers):
            self.answer_var.set(self.exam_answers[self.current_question_index])

    def finish_exam(self):
        self.save_current_answer()
        unanswered = [i+1 for i,a in enumerate(self.exam_answers) if not a]
        if unanswered:
            if not messagebox.askyesno("Sin responder",
                f"{len(unanswered)} preguntas sin responder. ¿Finalizar?"):
                return
        self.calculate_exam_results()
        self.show_exam_results()

    def calculate_exam_results(self):
        self.exam_results = []
        cc = 0
        for i,(q,a) in enumerate(zip(self.exam_questions, self.exam_answers)):
            ok = a == q['correcta']
            if ok: cc += 1
            self.exam_results.append({'question_num':i+1,'question':q,
                'user_answer':a,'correct_answer':q['correcta'],'is_correct':ok})
            self.update_statistics(q, ok)
        self.correct_count = cc
        self.total_questions = len(self.exam_questions)
        self.percentage = (cc/self.total_questions*100) if self.total_questions > 0 else 0

    def confirm_exit_exam(self):
        if messagebox.askyesno("Salir","¿Seguro? Se perderá el progreso."):
            self.show_home_screen()

    def start_single_topic_exam(self):
        self.show_topic_selection_screen(True)

    def start_multi_topic_exam(self):
        self.show_topic_selection_screen(False)

    def start_complete_exam(self):
        qs = self.load_questions()
        if not qs: messagebox.showerror("Error","No hay preguntas."); return
        self.show_exam_config_screen(qs, "Examen Completo")

    def start_failed_questions_exam(self):
        stats = self.load_statistics()
        if not stats: messagebox.showinfo("Info","Haz exámenes primero."); return
        failed = sorted([
            {'tema':s['tema'],'pregunta':s['pregunta'],
             'pct':(s['fallos']/s['intentos'])*100,'fallos':s['fallos']}
            for s in stats.values() if s['intentos']>=2 and s['fallos']>0
        ], key=lambda x: x['pct'], reverse=True)[:20]
        if not failed: messagebox.showinfo("Info","No hay suficientes datos."); return
        all_qs = self.load_questions()
        exam_qs = [q for f in failed for q in all_qs
                   if q['tema']==f['tema'] and q['pregunta']==f['pregunta']][:len(failed)]
        if not exam_qs: messagebox.showerror("Error","No se pudieron cargar."); return
        self.show_exam_config_screen(exam_qs, "Preguntas más Falladas")

    def show_topic_selection_screen(self, single=True):
        self.clear_content()
        ctk.CTkLabel(self.content_frame,
            text="Selecciona un tema" if single else "Selecciona temas",
            font=ctk.CTkFont(size=22, weight="bold")).pack(pady=(30,20))
        mf = ctk.CTkFrame(self.content_frame)
        mf.pack(fill="both", expand=True, padx=40, pady=20)
        lf = ctk.CTkFrame(mf)
        lf.pack(side="left", fill="both", expand=True, padx=(20,10), pady=20)
        rf = ctk.CTkFrame(mf, width=280)
        rf.pack(side="right", fill="y", padx=(10,20), pady=20)
        rf.pack_propagate(False)
        temas = self.get_available_topics()
        if not temas: messagebox.showerror("Error","No hay temas."); self.show_home_screen(); return
        ctk.CTkLabel(lf, text="Temas Disponibles",
            font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(20,15))
        tsf = ctk.CTkScrollableFrame(lf)
        tsf.pack(fill="both", expand=True, padx=20, pady=(0,20))
        if single:
            sv = ctk.StringVar()
            for t in sorted(temas):
                tf = ctk.CTkFrame(tsf)
                tf.pack(fill="x", pady=4, padx=10)
                ctk.CTkRadioButton(tf, text=t, variable=sv, value=t,
                    font=ctk.CTkFont(size=16),
                    command=lambda tt=t: self.update_topic_info(tt)).pack(side="left", padx=15, pady=12)
        else:
            svars = {t: tk.BooleanVar() for t in sorted(temas)}
            for t,v in svars.items():
                tf = ctk.CTkFrame(tsf)
                tf.pack(fill="x", pady=4, padx=10)
                ctk.CTkCheckBox(tf, text=t, variable=v, font=ctk.CTkFont(size=16)).pack(side="left", padx=15, pady=12)
        bf = ctk.CTkFrame(lf, fg_color="transparent")
        bf.pack(fill="x", padx=20, pady=(0,20))
        ctk.CTkButton(bf, text="Volver", command=self.show_home_screen, width=110, height=38).pack(side="left")
        if single:
            ctk.CTkButton(bf, text="Continuar",
                command=lambda: self.continue_single_topic(sv.get()),
                width=110, height=38).pack(side="right")
        else:
            ctk.CTkButton(bf, text="Continuar",
                command=lambda: self.continue_multi_topic(svars),
                width=110, height=38).pack(side="right")
        ctk.CTkLabel(rf, text="Info del Tema",
            font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(20,15))
        self.info_text_widget = ctk.CTkTextbox(rf, height=200, width=240)
        self.info_text_widget.pack(padx=15, pady=10)
        self.info_text_widget.insert("1.0","Selecciona un tema.")

    def update_topic_info(self, tema):
        if not hasattr(self,'info_text_widget'): return
        qs = self.load_questions(tema)
        imgs = sum(1 for q in qs if q.get('imagen'))
        stats = self.load_statistics()
        ts = {'intentos':0,'aciertos':0,'fallos':0}
        for s in stats.values():
            if s['tema']==tema:
                ts['intentos']+=s['intentos']; ts['aciertos']+=s['aciertos']; ts['fallos']+=s['fallos']
        info = f"Tema: {tema}\n\nPreguntas: {len(qs)}\nCon imagen: {imgs}\n\nEstadísticas:\nIntentos: {ts['intentos']}\nAciertos: {ts['aciertos']}\nFallos: {ts['fallos']}"
        if ts['intentos']>0: info += f"\nAcierto: {ts['aciertos']/ts['intentos']*100:.1f}%"
        self.info_text_widget.delete("1.0","end")
        self.info_text_widget.insert("1.0", info)

    def continue_single_topic(self, t):
        if not t: messagebox.showwarning("Aviso","Selecciona un tema."); return
        qs = self.load_questions(t)
        if not qs: messagebox.showerror("Error",f"Sin preguntas en {t}"); return
        self.show_exam_config_screen(qs, f"Examen: {t}")

    def continue_multi_topic(self, svars):
        sel = [t for t,v in svars.items() if v.get()]
        if not sel: messagebox.showwarning("Aviso","Selecciona al menos un tema."); return
        qs = []
        for t in sel: qs.extend(self.load_questions(t))
        if not qs: messagebox.showerror("Error","Sin preguntas."); return
        self.show_exam_config_screen(qs, f"Multi-tema: {', '.join(sorted(sel))}")

    def show_exam_config_screen(self, questions, title):
        self.clear_content()
        ctk.CTkLabel(self.content_frame, text="Configuración del Examen",
            font=ctk.CTkFont(size=22, weight="bold")).pack(pady=(30,20))
        cf = ctk.CTkFrame(self.content_frame)
        cf.pack(fill="both", expand=True, padx=40, pady=20)
        ctk.CTkLabel(cf, text=title, font=ctk.CTkFont(size=18, weight="bold")).pack(pady=15)
        ctk.CTkLabel(cf, text=f"Preguntas disponibles: {len(questions)}",
            font=ctk.CTkFont(size=16)).pack(pady=5)
        ctk.CTkLabel(cf, text="Número de preguntas:",
            font=ctk.CTkFont(size=16)).pack(pady=(20,10))
        self.num_questions_var = tk.IntVar(value=min(10,len(questions)))
        nf = ctk.CTkFrame(cf, fg_color="transparent")
        nf.pack(pady=10)
        ctk.CTkSlider(nf, from_=1, to=len(questions), variable=self.num_questions_var,
            number_of_steps=len(questions)-1 if len(questions)>1 else 1,
            width=280).pack(side="left", padx=(0,15))
        ql = ctk.CTkLabel(nf, text=str(self.num_questions_var.get()),
            font=ctk.CTkFont(size=15, weight="bold"))
        ql.pack(side="left")
        self.num_questions_var.trace('w', lambda *a: ql.configure(text=str(self.num_questions_var.get())))
        self.shuffle_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(cf, text="Mezclar preguntas", variable=self.shuffle_var).pack(pady=5)
        self.show_results_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(cf, text="Mostrar resultados al final", variable=self.show_results_var).pack(pady=5)
        bf = ctk.CTkFrame(cf, fg_color="transparent")
        bf.pack(fill="x", padx=20, pady=(20,20))
        ctk.CTkButton(bf, text="Volver", command=self.show_home_screen, width=110, height=38).pack(side="left")
        ctk.CTkButton(bf, text="Iniciar Examen",
            command=lambda: self.start_exam(questions, title),
            width=140, height=38, font=ctk.CTkFont(size=15, weight="bold")).pack(side="right")

    def start_exam(self, questions, title):
        n = self.num_questions_var.get()
        sel = questions.copy() if n>=len(questions) else random.sample(questions, n)
        if self.shuffle_var.get(): random.shuffle(sel)
        self.exam_questions = sel
        self.current_question_index = 0
        self.exam_answers = []
        self.exam_results = []
        self.exam_title = title
        self.show_results_at_end = self.show_results_var.get()
        self.show_question_screen()

    def show_exam_results(self):
        self.clear_content()
        ctk.CTkLabel(self.content_frame, text="Resultados del Examen",
            font=ctk.CTkFont(size=26, weight="bold")).pack(pady=(30,20))
        rf = ctk.CTkFrame(self.content_frame)
        rf.pack(fill="both", expand=True, padx=20, pady=(0,20))
        ctk.CTkLabel(rf, text=self.exam_title,
            font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(15,5))
        sf = ctk.CTkFrame(rf, fg_color="transparent")
        sf.pack(pady=10)
        for i,(lbl,val) in enumerate([
            ("Preguntas", str(self.total_questions)),
            ("Correctas", str(self.correct_count)),
            ("Incorrectas", str(self.total_questions-self.correct_count)),
            ("Porcentaje", f"{self.percentage:.1f}%")
        ]):
            c = ctk.CTkFrame(sf)
            c.grid(row=0, column=i, padx=8, pady=8, sticky="ew")
            ctk.CTkLabel(c, text=lbl, font=ctk.CTkFont(size=12), text_color="gray60").pack(pady=(8,3))
            ctk.CTkLabel(c, text=val, font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0,8))
        sf.grid_columnconfigure((0,1,2,3), weight=1)
        msgs = [("90","EXCELENTE! Campeón!","#A3BE8C"),("80","MUY BIEN! Gran trabajo","#88C999"),
                ("60","Buen trabajo, puedes mejorar","#EBCB8B"),("0","Necesitas estudiar más","#D08770")]
        for pct,msg,color in msgs:
            if self.percentage >= int(pct):
                ctk.CTkLabel(rf, text=msg, font=ctk.CTkFont(size=15, weight="bold"),
                    text_color=color).pack(pady=10); break
        df = ctk.CTkFrame(rf)
        df.pack(fill="both", expand=True, padx=20, pady=(0,20))
        ctk.CTkLabel(df, text="Detalle de Respuestas",
            font=ctk.CTkFont(size=15, weight="bold")).pack(pady=(12,8))
        ds = ctk.CTkScrollableFrame(df)
        ds.pack(fill="both", expand=True, padx=12, pady=(0,12))
        for r in self.exam_results:
            qf = ctk.CTkFrame(ds)
            qf.pack(fill="x", pady=4, padx=8)
            hf = ctk.CTkFrame(qf, fg_color="transparent")
            hf.pack(fill="x", padx=12, pady=(8,4))
            ctk.CTkLabel(hf, text=f"P{r['question_num']}",
                font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
            sc = "#A3BE8C" if r['is_correct'] else "#BF616A"
            ctk.CTkLabel(hf, text="Correcto" if r['is_correct'] else "Incorrecto",
                font=ctk.CTkFont(size=12), text_color=sc).pack(side="right")
            qt = r['question']['pregunta'][:80]+("..." if len(r['question']['pregunta'])>80 else "")
            ctk.CTkLabel(qf, text=qt, font=ctk.CTkFont(size=11), wraplength=650).pack(anchor="w", padx=12, pady=(0,4))
            af = ctk.CTkFrame(qf, fg_color="transparent")
            af.pack(fill="x", padx=12, pady=(0,8))
            uc = "#A3BE8C" if r['is_correct'] else "#BF616A"
            ctk.CTkLabel(af, text=f"Tu respuesta: {r['user_answer'] or 'Sin respuesta'}",
                font=ctk.CTkFont(size=11), text_color=uc).pack(anchor="w")
            if not r['is_correct']:
                ctk.CTkLabel(af, text=f"Correcta: {r['correct_answer']}",
                    font=ctk.CTkFont(size=11), text_color="#A3BE8C").pack(anchor="w")
        bfr = ctk.CTkFrame(rf, fg_color="transparent")
        bfr.pack(fill="x", padx=20, pady=(0,20))
        ctk.CTkButton(bfr, text="Inicio", command=self.show_home_screen, width=130, height=38).pack(side="left")
        ctk.CTkButton(bfr, text="Estadísticas", command=self.show_statistics_screen, width=130, height=38).pack(side="right")
        ctk.CTkButton(bfr, text="Repetir", command=self.repeat_exam, width=130, height=38).pack(side="right", padx=(0,10))

    def repeat_exam(self):
        self.current_question_index = 0
        self.exam_answers = []
        self.exam_results = []
        if hasattr(self,'shuffle_var') and self.shuffle_var.get():
            random.shuffle(self.exam_questions)
        self.show_question_screen()

    def load_statistics(self):
        if Path(self.estadisticas_file).exists():
            try:
                df = pd.read_excel(self.estadisticas_file)
                return {f"{r['tema']}:{r['pregunta_corta']}": {
                    'tema':r['tema'],'pregunta':r['pregunta_completa'],
                    'intentos':int(r['intentos']),'aciertos':int(r['aciertos']),'fallos':int(r['fallos'])
                } for _,r in df.iterrows()}
            except: return {}
        return {}

    def save_statistics(self, stats):
        try:
            datos = [{'tema':s['tema'],'pregunta_corta':s['pregunta'][:50].strip(),
                'pregunta_completa':s['pregunta'],'intentos':s['intentos'],
                'aciertos':s['aciertos'],'fallos':s['fallos'],
                'pct_fallos':round(s['fallos']/s['intentos']*100,1) if s['intentos']>0 else 0,
                'pct_aciertos':round(s['aciertos']/s['intentos']*100,1) if s['intentos']>0 else 0}
                for s in stats.values()]
            df = pd.DataFrame(datos).sort_values(['pct_fallos','fallos'],ascending=[False,False])
            with pd.ExcelWriter(self.estadisticas_file, engine='openpyxl') as w:
                df.to_excel(w, sheet_name='Estadisticas', index=False)
                ws = w.sheets['Estadisticas']
                for col,width in {'A':20,'B':50,'C':80,'D':12,'E':12,'F':12,'G':18,'H':20}.items():
                    ws.column_dimensions[col].width = width
        except Exception as e: print(f"Error stats: {e}")

    def update_statistics(self, q, ok):
        stats = self.load_statistics()
        k = f"{q['tema']}:{q['pregunta'][:50].strip()}"
        if k not in stats:
            stats[k] = {'tema':q['tema'],'pregunta':q['pregunta'],'intentos':0,'aciertos':0,'fallos':0}
        stats[k]['intentos'] += 1
        if ok: stats[k]['aciertos'] += 1
        else: stats[k]['fallos'] += 1
        self.save_statistics(stats)

    def get_most_failed_questions(self, n=10):
        stats = self.load_statistics()
        failed = [{'tema':s['tema'],'pregunta':s['pregunta'],'fallos':s['fallos'],
            'intentos':s['intentos'],'pct':s['fallos']/s['intentos']*100}
            for s in stats.values() if s['intentos']>=2 and s['fallos']>0]
        return sorted(failed, key=lambda x: (x['pct'],x['fallos']), reverse=True)[:n]

    def show_statistics_screen(self):
        self.clear_content()
        ctk.CTkLabel(self.content_frame, text="Estadísticas de Rendimiento",
            font=ctk.CTkFont(size=22, weight="bold")).pack(pady=(30,20))
        stats = self.load_statistics()
        if not stats:
            ef = ctk.CTkFrame(self.content_frame)
            ef.pack(fill="both", expand=True, padx=40, pady=20)
            ctk.CTkLabel(ef, text="Sin estadísticas aún.",
                font=ctk.CTkFont(size=16)).pack(pady=100)
            ctk.CTkButton(ef, text="Volver", command=self.show_home_screen, width=120, height=38).pack(pady=20)
            return
        sf = ctk.CTkFrame(self.content_frame)
        sf.pack(fill="both", expand=True, padx=20, pady=(0,20))
        ta = sum(s['intentos'] for s in stats.values())
        tc = sum(s['aciertos'] for s in stats.values())
        tf = sum(s['fallos'] for s in stats.values())
        op = (tc/ta*100) if ta>0 else 0
        gsf = ctk.CTkFrame(sf, fg_color="transparent")
        gsf.pack(pady=15)
        for i,(lbl,val) in enumerate([
            ("Intentos",str(ta)),("Aciertos",str(tc)),("Fallos",str(tf)),("Promedio",f"{op:.1f}%")
        ]):
            c = ctk.CTkFrame(gsf)
            c.grid(row=0,column=i,padx=6,pady=6,sticky="ew")
            ctk.CTkLabel(c,text=lbl,font=ctk.CTkFont(size=11),text_color="gray60").pack(pady=(8,3))
            ctk.CTkLabel(c,text=val,font=ctk.CTkFont(size=15,weight="bold")).pack(pady=(0,8))
        gsf.grid_columnconfigure((0,1,2,3),weight=1)
        ctk.CTkLabel(sf,text="Preguntas más Falladas",
            font=ctk.CTkFont(size=16,weight="bold")).pack(pady=(10,8))
        fqs = self.get_most_failed_questions(5)
        if fqs:
            fsc = ctk.CTkScrollableFrame(sf, height=180)
            fsc.pack(fill="both", expand=True, padx=20, pady=(0,15))
            for i,fq in enumerate(fqs,1):
                ff = ctk.CTkFrame(fsc)
                ff.pack(fill="x", pady=4, padx=8)
                ctk.CTkLabel(ff, text=f"{i}. {fq['pregunta'][:65]}...",
                    font=ctk.CTkFont(size=12), wraplength=580).pack(anchor="w", padx=10, pady=(8,4))
                ctk.CTkLabel(ff, text=f"{fq['fallos']} fallos / {fq['intentos']} intentos ({fq['pct']:.1f}%)",
                    font=ctk.CTkFont(size=11), text_color="#BF616A").pack(anchor="w", padx=10, pady=(0,8))
        ctk.CTkButton(sf, text="Volver", command=self.show_home_screen,
            width=120, height=38).pack(side="bottom", padx=20, pady=(0,20))

    def show_settings_screen(self):
        self.clear_content()
        ctk.CTkLabel(self.content_frame, text="Configuración",
            font=ctk.CTkFont(size=22, weight="bold")).pack(pady=(30,20))
        sf = ctk.CTkFrame(self.content_frame)
        sf.pack(fill="both", expand=True, padx=40, pady=20)
        dp = Path('datos'); ip = Path('imagenes')
        archivos = len(list(dp.glob('*.xlsx'))+list(dp.glob('*.csv'))) if dp.exists() else 0
        imagenes = len(list(ip.glob('*'))) if ip.exists() else 0
        info = (f"Carpeta datos: {dp.absolute()}\n"
                f"Archivos de preguntas: {archivos}\n\n"
                f"Carpeta imágenes: {ip.absolute()}\n"
                f"Imágenes encontradas: {imagenes}")
        ctk.CTkLabel(sf, text=info, font=ctk.CTkFont(size=12), justify="left").pack(anchor="w", padx=20, pady=20)
        bf = ctk.CTkFrame(sf, fg_color="transparent")
        bf.pack(fill="x", padx=20, pady=10)
        ctk.CTkButton(bf, text="Abrir Carpeta Datos",
            command=lambda: self._open_folder(Path('datos')), width=160, height=35).pack(side="left", padx=(0,10))
        ctk.CTkButton(bf, text="Abrir Carpeta Imágenes",
            command=lambda: self._open_folder(Path('imagenes')), width=170, height=35).pack(side="left", padx=10)
        ctk.CTkButton(sf, text="Volver", command=self.show_home_screen,
            width=110, height=38).pack(side="bottom", padx=20, pady=(0,20))

    def _open_folder(self, path):
        if not path.exists():
            try: path.mkdir(); messagebox.showinfo("OK",f"Creada: {path}")
            except Exception as e: messagebox.showerror("Error",str(e)); return
        try:
            import sys
            os.system(f'open "{path}"' if sys.platform=='darwin' else f'xdg-open "{path}"')
        except Exception as e: messagebox.showerror("Error",str(e))

    def run(self):
        for f in ['datos','imagenes']:
            Path(f).mkdir(exist_ok=True)
        self.root.mainloop()


def main():
    for m in ['customtkinter','pandas','openpyxl','PIL']:
        try: __import__(m)
        except ImportError:
            print(f"Falta: {'pillow' if m=='PIL' else m}")
            print(f"Instala con: pip3 install {'pillow' if m=='PIL' else m}")
            input("Enter para salir..."); return
    ExamSystemGUI().run()

if __name__ == "__main__":
    main()
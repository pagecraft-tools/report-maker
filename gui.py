# -*- coding: utf-8 -*-
"""보고서 제작기 창.
  - 새 보고서 만들기   : 원본(엑셀·워드·파워포인트·한글·PDF·붙여넣은 글) → 보고서
  - 기존 보고서에 추가 : 만든 보고서 + 추가 원고 → 같은 양식으로 이어 붙인 새 파일
예전에 만든 보고서(샘플)를 넣으면 그 모양 그대로 만든다. AI 없이 프로그램만으로 동작.
원고의 글자·숫자는 바꾸지 않고, 끝나면 원본 대조 결과를 보여 준다.
"""
import io
import json
import os
import queue
import subprocess
import sys
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
SETTINGS = os.path.join(HERE, 'settings.json')
FILETYPES = [('보고서·원고 파일', '*.xlsx *.xlsm *.docx *.pptx *.hwpx *.hwp *.pdf *.txt'), ('모든 파일', '*.*')]
SAMPLE_TYPES = [('워드·엑셀 보고서', '*.docx *.xlsx *.xlsm'), ('모든 파일', '*.*')]


def load_settings():
    try:
        with io.open(SETTINGS, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(d):
    try:
        with io.open(SETTINGS, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


class QueueWriter:
    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(('log', s))

    def flush(self):
        pass


def open_path(p):
    try:
        os.startfile(p)
    except Exception:
        pass


def main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import engine

    BG, CARD, LINE = '#F4F6F9', '#FFFFFF', '#DDE2E8'
    INK, GRAY, SOFT = '#1E2430', '#7A808A', '#A4AAB3'
    ACC, ACC_H, ACC_SOFT = '#3B6EA5', '#2F5C8C', '#EAF1F8'
    OK, BAD = '#2E7D4F', '#C0392B'
    F = '맑은 고딕'

    st = load_settings()
    root = tk.Tk()
    root.title('보고서 제작기')
    root.configure(bg=BG)
    k = max(1.0, root.winfo_fpixels('1i') / 96.0)
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    W, H = min(int(820 * k), sw - 40), min(int(900 * k), sh - 80)
    root.geometry(f'{W}x{H}+{(sw - W) // 2}+{max(0, (sh - H) // 2 - 20)}')
    root.minsize(min(int(700 * k), W), min(int(640 * k), H))
    try:
        root.iconbitmap(os.path.join(HERE, 'icon.ico'))
    except Exception:
        pass
    sty = ttk.Style(root)
    try:
        sty.theme_use('clam')
    except Exception:
        pass
    sty.configure('Vertical.TScrollbar', background=LINE, troughcolor=CARD, bordercolor=CARD, arrowcolor=GRAY)

    # 스크롤되는 본문
    outer = tk.Frame(root, bg=BG)
    outer.pack(fill='both', expand=True)
    canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
    vs = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
    canvas.configure(yscrollcommand=vs.set)
    vs.pack(side='right', fill='y')
    canvas.pack(side='left', fill='both', expand=True)
    page = tk.Frame(canvas, bg=BG)
    win = canvas.create_window((0, 0), window=page, anchor='nw')
    page.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
    canvas.bind('<Configure>', lambda e: canvas.itemconfigure(win, width=e.width))
    canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(int(-e.delta / 120), 'units')
                    if not isinstance(e.widget, tk.Text) else None)

    def card(parent, title, hint=''):
        o = tk.Frame(parent, bg=LINE)
        o.pack(fill='x', padx=24, pady=(0, 14))
        box = tk.Frame(o, bg=CARD, padx=20, pady=16)
        box.pack(fill='both', expand=True, padx=1, pady=1)
        head = tk.Frame(box, bg=CARD)
        head.pack(fill='x', pady=(0, 10))
        lab = tk.Label(head, text=title, font=(F, 12, 'bold'), fg=INK, bg=CARD)
        lab.pack(side='left')
        hl = tk.Label(head, text=hint, font=(F, 9), fg=GRAY, bg=CARD)
        hl.pack(side='left', padx=(10, 0), pady=(3, 0))
        return o, box, lab, hl

    def button(parent, text, cmd, primary=False, small=False):
        if primary:
            return tk.Button(parent, text=text, command=cmd, font=(F, 11, 'bold'), fg='white', bg=ACC,
                             activebackground=ACC_H, activeforeground='white', relief='flat', bd=0,
                             padx=26, pady=9, cursor='hand2', disabledforeground='#D5DEE9')
        return tk.Button(parent, text=text, command=cmd, font=(F, 9 if small else 10), fg=ACC, bg=ACC_SOFT,
                         activebackground='#DCE7F3', activeforeground=ACC_H, relief='flat', bd=0,
                         padx=12 if small else 16, pady=5 if small else 7, cursor='hand2', disabledforeground=SOFT)

    def entry(parent, var):
        f = tk.Frame(parent, bg=LINE)
        e = tk.Entry(f, textvariable=var, font=(F, 10), relief='flat', bd=0, bg='#FBFCFD', fg=INK)
        e.pack(fill='x', padx=1, pady=1, ipady=6, ipadx=6)
        return f

    # ── 헤더
    hdr = tk.Frame(page, bg=BG)
    hdr.pack(fill='x', padx=24, pady=(20, 12))
    tk.Label(hdr, text='보고서 제작기', font=(F, 17, 'bold'), fg=INK, bg=BG).pack(anchor='w')
    tk.Label(hdr, text='예전에 만든 보고서를 넣으면, 새 원고를 그 모양 그대로 만들어 줍니다.\n'
                       '원고의 글자·숫자는 절대 바꾸지 않으며, 끝나면 원본과 한 칸씩 대조한 결과를 보여 줍니다.',
             font=(F, 10), fg=GRAY, bg=BG, justify='left').pack(anchor='w', pady=(2, 0))

    # ── 모드
    mode = tk.StringVar(value=st.get('mode', 'add'))
    mrow = tk.Frame(page, bg=BG)
    mrow.pack(fill='x', padx=24, pady=(0, 14))
    mbtn = {}

    def set_mode(m):
        mode.set(m)
        for key, b in mbtn.items():
            on = key == m
            b.configure(bg=ACC if on else CARD, fg='white' if on else INK,
                        activebackground=ACC_H if on else ACC_SOFT)
        if m == 'add':
            new_box_o.pack_forget()
            add_box_o.pack(fill='x', padx=24, pady=(0, 14), before=src_o)
            src_lab.configure(text='② 추가할 원고')
            src_hint.configure(text='이번에 새로 넣을 원고 (예: 12호 발송결과 엑셀) — 여러 개면 호마다 하나씩 붙습니다')
        else:
            add_box_o.pack_forget()
            new_box_o.pack(fill='x', padx=24, pady=(0, 14), before=src_o)
            src_lab.configure(text='② 원본 파일')
            src_hint.configure(text='보고서에 들어갈 원본 (엑셀·워드·파워포인트·한글·PDF) — 파일 하나가 한 묶음(호)')
        suggest_out()

    for key, label in (('add', '기존 보고서에 추가'), ('new', '새 보고서 만들기')):
        b = tk.Button(mrow, text=label, command=lambda kk=key: set_mode(kk), font=(F, 11, 'bold'), relief='flat',
                      bd=0, padx=22, pady=9, cursor='hand2')
        b.pack(side='left', padx=(0, 8))
        mbtn[key] = b

    # ── ① 기존 보고서 (추가 모드)
    add_box_o, add_box, _, _ = card(page, '① 기존 보고서', '지금까지 만든 보고서 (워드 .docx 또는 엑셀)')
    base_var = tk.StringVar(value='')
    r = tk.Frame(add_box, bg=CARD)
    r.pack(fill='x')
    entry(r, base_var).pack(side='left', fill='x', expand=True)

    def pick_base():
        p = filedialog.askopenfilename(title='기존 보고서 고르기', filetypes=SAMPLE_TYPES,
                                       initialdir=st.get('last_dir') or os.path.expanduser('~'))
        if p:
            base_var.set(os.path.normpath(p))
            st['last_dir'] = os.path.dirname(p)
            suggest_out()
    button(r, '고르기…', pick_base, small=True).pack(side='left', padx=(8, 0))
    tk.Label(add_box, text='기존 파일은 그대로 두고, 추가분을 넣은 새 파일을 따로 만듭니다 (예: …_1-11호 → …_1-12호).\n'
                           '마지막 호의 모양을 그대로 복제해 새 호를 만듭니다. 한글(.hwp) 보고서는 한글에서\n'
                           '[다른 이름으로 저장 → 파일 형식: 워드 문서(*.docx)]로 바꿔서 넣어 주세요.',
             font=(F, 9), fg=GRAY, bg=CARD, justify='left').pack(anchor='w', pady=(6, 0))

    # ── ① 새 보고서 설정 (새로 만들기 모드)
    new_box_o, new_box, _, _ = card(page, '① 보고서 설정', '비워 둬도 됩니다')
    title_var = tk.StringVar(value='')
    tmpl_var = tk.StringVar(value='')
    tk.Label(new_box, text='보고서 제목 (파일 이름)', font=(F, 9), fg=GRAY, bg=CARD).pack(anchor='w', pady=(0, 4))
    entry(new_box, title_var).pack(fill='x')
    tk.Label(new_box, text='샘플 보고서 — 예전에 만든 보고서(.docx·엑셀)를 넣으면 그 모양 그대로 만듭니다 (없으면 기본 모양)',
             font=(F, 9), fg=GRAY, bg=CARD).pack(anchor='w', pady=(10, 4))
    r2 = tk.Frame(new_box, bg=CARD)
    r2.pack(fill='x')
    entry(r2, tmpl_var).pack(side='left', fill='x', expand=True)

    def pick_tmpl():
        p = filedialog.askopenfilename(title='샘플 보고서 고르기', filetypes=SAMPLE_TYPES,
                                       initialdir=st.get('last_dir') or os.path.expanduser('~'))
        if p:
            tmpl_var.set(os.path.normpath(p))
    button(r2, '고르기…', pick_tmpl, small=True).pack(side='left', padx=(8, 0))
    button(r2, '비우기', lambda: tmpl_var.set(''), small=True).pack(side='left', padx=(6, 0))

    # ── ② 원본 / 추가 원고
    src_o, src_box, src_lab, src_hint = card(page, '② 추가할 원고')
    lw = tk.Frame(src_box, bg=LINE)
    lw.pack(fill='x')
    lst = tk.Listbox(lw, height=5, font=(F, 10), relief='flat', bd=0, bg='#FBFCFD', fg=INK,
                     selectbackground=ACC, activestyle='none', highlightthickness=0)
    lst.pack(fill='x', padx=1, pady=1)
    files = []

    def refresh():
        lst.delete(0, 'end')
        for p in files:
            lst.insert('end', '  ' + os.path.basename(p) + '   —   ' + os.path.dirname(p))
        suggest_out()

    def add_files():
        ps = filedialog.askopenfilenames(title='원고 파일 고르기', filetypes=FILETYPES,
                                         initialdir=st.get('last_dir') or os.path.expanduser('~'))
        for p in ps:
            p = os.path.normpath(p)
            if p not in files:
                files.append(p)
        if ps:
            st['last_dir'] = os.path.dirname(ps[0])
        refresh()

    def del_sel():
        for i in reversed(lst.curselection()):
            files.pop(i)
        refresh()

    def move(d):
        sel = lst.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + d
        if 0 <= j < len(files):
            files[i], files[j] = files[j], files[i]
            refresh()
            lst.selection_set(j)

    br = tk.Frame(src_box, bg=CARD)
    br.pack(fill='x', pady=(8, 0))
    button(br, '파일 추가…', add_files, small=True).pack(side='left')
    button(br, '선택 빼기', del_sel, small=True).pack(side='left', padx=(6, 0))
    button(br, '▲', lambda: move(-1), small=True).pack(side='left', padx=(6, 0))
    button(br, '▼', lambda: move(1), small=True).pack(side='left', padx=(4, 0))
    tk.Label(br, text='위에서부터 순서대로 들어갑니다', font=(F, 9), fg=GRAY, bg=CARD).pack(side='left', padx=(10, 0))

    tk.Label(src_box, text='복사한 글 붙여넣기 (선택) — 메일·카톡·웹에서 복사한 글을 그대로 붙여넣으세요',
             font=(F, 9), fg=GRAY, bg=CARD).pack(anchor='w', pady=(12, 4))
    tw = tk.Frame(src_box, bg=LINE)
    tw.pack(fill='x')
    paste = tk.Text(tw, height=5, font=(F, 10), wrap='word', relief='flat', bd=0, bg='#FBFCFD', fg=INK,
                    padx=8, pady=6, undo=True)
    paste.pack(fill='x', padx=1, pady=1)

    # ── 매호 이미지 (선택)
    oi, bi, _, _ = card(page, '매호 이미지 폴더 (선택)', '비워 두면 이미지 없이 만듭니다')
    img_var = tk.StringVar(value=st.get('img_dir', ''))
    ri = tk.Frame(bi, bg=CARD)
    ri.pack(fill='x')
    entry(ri, img_var).pack(side='left', fill='x', expand=True)

    def pick_img():
        p = filedialog.askdirectory(title='매호 이미지 폴더 고르기',
                                    initialdir=img_var.get() or st.get('last_dir') or '')
        if p:
            img_var.set(os.path.normpath(p))
    button(ri, '고르기…', pick_img, small=True).pack(side='left', padx=(8, 0))
    button(ri, '비우기', lambda: img_var.set(''), small=True).pack(side='left', padx=(6, 0))
    button(ri, '열기', lambda: open_path(img_var.get()) if img_var.get() else None,
           small=True).pack(side='left', padx=(6, 0))
    tk.Label(bi, text='파일 이름에 호 번호를 넣어 저장하세요: 26_12호.png → 12호에 들어갑니다 (앞의 26_은 연도).\n'
                      '한 호에 여러 장이면 26_12호_1.png, 26_12호_2.png 순서대로. PNG·JPG 모두 됩니다.\n'
                      '가로 2000px보다 크면 자동으로 줄이고, 지난 호에 그림이 있던 자리에 같은 폭으로 넣습니다\n'
                      '(그림 자리가 없으면 호 제목 바로 아래에 본문 폭으로).\n'
                      '★ 이미지만 넣기: 「기존 보고서에 추가」에서 ② 원고를 비워 두고 이 폴더만 고르면,\n'
                      '   보고서에 이미 있는 호마다 그 호 그림을 넣은 새 파일(…_이미지.docx)을 만듭니다.',
             font=(F, 9), fg=GRAY, bg=CARD, justify='left').pack(anchor='w', pady=(6, 0))

    # ── ③ 요청 사항·저장 위치
    o3, b3, _, _ = card(page, '③ 저장 폴더', '처음엔 기존 보고서(또는 첫 원고)가 있는 폴더')
    out_var = tk.StringVar(value='')
    r3 = tk.Frame(b3, bg=CARD)
    r3.pack(fill='x')
    entry(r3, out_var).pack(side='left', fill='x', expand=True)
    out_manual = [False]

    def pick_out():
        p = filedialog.askdirectory(title='저장 폴더 고르기', initialdir=out_var.get() or st.get('last_dir') or '')
        if p:
            out_var.set(os.path.normpath(p))
            out_manual[0] = True
    button(r3, '바꾸기…', pick_out, small=True).pack(side='left', padx=(8, 0))

    def suggest_out():
        if out_manual[0]:
            return
        ref = base_var.get() if mode.get() == 'add' and base_var.get() else (files[0] if files else '')
        if ref:
            out_var.set(os.path.dirname(ref))

    # ── 실행
    run_row = tk.Frame(page, bg=BG)
    run_row.pack(fill='x', padx=24, pady=(0, 10))
    go = button(run_row, '만들기', lambda: start(), primary=True)
    go.pack(side='left')
    status = tk.Label(run_row, text='', font=(F, 10, 'bold'), fg=GRAY, bg=BG)
    status.pack(side='left', padx=(14, 0))

    res_row = tk.Frame(page, bg=BG)
    res_row.pack(fill='x', padx=24, pady=(0, 10))
    result = {'out': None, 'rep': None}
    b_open = button(res_row, '결과 파일 열기', lambda: open_path(result['out']))
    b_rep = button(res_row, '원본 대조 결과 보기', lambda: open_path(result['rep']))
    b_dir = button(res_row, '폴더 열기', lambda: open_path(os.path.dirname(result['out'])))

    lo, lbox, _, _ = card(page, '진행 기록')
    lw2 = tk.Frame(lbox, bg=LINE)
    lw2.pack(fill='x')
    logt = tk.Text(lw2, height=8, font=(F, 9), wrap='word', relief='flat', bd=0, bg='#FBFCFD', fg=INK,
                   padx=8, pady=6, state='disabled')
    logt.pack(fill='x', padx=1, pady=1)


    q = queue.Queue()

    def put_log(s):
        logt.configure(state='normal')
        logt.insert('end', s)
        logt.see('end')
        logt.configure(state='disabled')
        for line in s.splitlines():
            if line.startswith('== '):
                status.configure(text=line[3:], fg=ACC)

    def poll():
        try:
            while True:
                kind, v = q.get_nowait()
                if kind == 'log':
                    put_log(v)
                elif kind == 'done':
                    out, rep, ok = v
                    result.update(out=out, rep=rep)
                    go.configure(state='normal')
                    status.configure(text='✔ 완료 — 원본과 글자·숫자가 모두 같습니다' if ok else
                                     '⚠ 완료 — 원본과 다른 곳이 있습니다. 대조 결과를 꼭 확인하세요', fg=OK if ok else BAD)
                    put_log(f'\n결과: {out}\n대조: {rep}\n')
                    for b in (b_open, b_rep, b_dir):
                        b.pack(side='left', padx=(0, 8))
                    if not ok:
                        open_path(rep)
                elif kind == 'err':
                    go.configure(state='normal')
                    status.configure(text='✘ 실패', fg=BAD)
                    put_log('\n' + v + '\n')
                    messagebox.showerror('보고서 제작기', v)
        except queue.Empty:
            pass
        root.after(150, poll)

    def start():
        m = mode.get()
        txt = paste.get('1.0', 'end').strip()
        if m == 'add' and not base_var.get():
            messagebox.showwarning('보고서 제작기', '① 기존 보고서를 골라 주세요.')
            return
        only_img = m == 'add' and not files and not txt and img_var.get().strip()
        if not files and not txt and not only_img:
            messagebox.showwarning('보고서 제작기', '② 원고 파일을 넣거나 글을 붙여넣어 주세요.\n'
                                   '(기존 보고서에 이미지만 넣으려면 원고는 비워 두고 매호 이미지 폴더만 고르세요)')
            return
        out_dir = out_var.get().strip()
        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showwarning('보고서 제작기', '③ 저장 폴더를 골라 주세요.')
            return
        img_dir = img_var.get().strip()
        if img_dir and not os.path.isdir(img_dir):
            messagebox.showwarning('보고서 제작기', '매호 이미지 폴더를 찾을 수 없습니다. 다시 골라 주세요.')
            return
        st.update(mode=m, img_dir=img_dir)
        save_settings(st)
        go.configure(state='disabled')
        for b in (b_open, b_rep, b_dir):
            b.pack_forget()
        logt.configure(state='normal')
        logt.delete('1.0', 'end')
        logt.configure(state='disabled')
        status.configure(text='시작…', fg=ACC)
        args = dict(m=m, base=base_var.get(), files=list(files), txt=txt,
                    out_dir=out_dir, title=title_var.get().strip(), tmpl=tmpl_var.get().strip(), img=img_dir,
                    only_img=bool(only_img))

        def work():
            old = sys.stdout
            sys.stdout = QueueWriter(q)
            try:
                if args['only_img']:
                    r = engine.add_images(args['base'], args['out_dir'], args['img'])
                elif args['m'] == 'add':
                    r = engine.add(args['base'], args['files'], args['out_dir'], paste=args['txt'],
                                   img_dir=args['img'])
                else:
                    r = engine.new(args['files'], args['out_dir'], args['title'], args['tmpl'], paste=args['txt'],
                                   img_dir=args['img'])
                q.put(('done', r))
            except Exception as e:
                q.put(('err', str(e)))
            finally:
                sys.stdout = old
        threading.Thread(target=work, daemon=True).start()

    set_mode(mode.get())
    root.after(150, poll)
    root.mainloop()


if __name__ == '__main__':
    main()

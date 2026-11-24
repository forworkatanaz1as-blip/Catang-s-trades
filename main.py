# -*- coding: utf-8 -*-
"""Android/Kivy-перенос генератора торговцев One Piece."""
import os, json, random, shutil
from kivy.app import App
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.properties import StringProperty, ListProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.spinner import Spinner
from kivy.uix.checkbox import CheckBox
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from openpyxl import load_workbook

RARITIES = ['Обычный', 'Необычный', 'Редкий', 'Очень редкий']
RARITY_DISTRIBUTION = {'Обычный':[1,0,0,0], 'Необычный':[.8,.2,0,0], 'Редкий':[.2,.6,.2,0], 'Очень редкий':[.3,.3,.3,.1]}
CAPITAL_RANGES = {'Обычный':(1000,10000),'Необычный':(10000,100000),'Редкий':(100000,1000000),'Очень редкий':(1000000,10000000)}
SPECS_BY_RARITY = {'Обычный':['торгаш','бакалейщик','овощной','мясник','рыбак','разливщик','кок','плотник','ткач'], 'Необычный':['литейщик','стеклодув','ботаник','доктор','бронник','клинковщик','стрелок','навигатор'], 'Редкий':['ювелир','полимат','антиквар','музыкант','книжник'], 'Очень редкий':['корабельщик','зверолов','делец']}
ALL_SPECS = ['Случайно'] + sorted({s for v in SPECS_BY_RARITY.values() for s in v})
NAMES = ['Арден','Борис','Валерий','Гаррет','Дариус','Елена','Жанна','Зорян','Кай','Лира','Марк','Нора','Оскар','Петра','Рен','Саймон','Талия','Ульрик','Фиона','Акира','Сакура','Роберт','Иван','Никита','Алекс','Майя','Юки','Рэн']

KV = '''
<TabbedPanel>:
    do_default_tab: False
    tab_width: dp(150)
    TabbedPanelItem:
        text: 'Генератор'
        GeneratorScreen:
    TabbedPanelItem:
        text: 'Библиотека'
        LibraryScreen:
    TabbedPanelItem:
        text: 'База товаров'
        DatabaseScreen:
<GeneratorScreen>:
    orientation: 'vertical'
    padding: dp(8); spacing: dp(6)
    BoxLayout:
        size_hint_y: None; height: dp(42); spacing: dp(5)
        Spinner: id: sr; text: 'Случайно'; values: ['Случайно'] + ['Обычный','Необычный','Редкий','Очень редкий']; size_hint_x: .5
        Spinner: id: seller; text: 'Случайно'; values: ['Случайно'] + ['Обычный','Необычный','Редкий','Очень редкий']; size_hint_x: .5
    BoxLayout:
        size_hint_y: None; height: dp(42); spacing: dp(5)
        Spinner: id: spec; text: 'Случайно'; values: app.spec_values
        Spinner: id: assort; text: 'Случайно'; values: ['Случайно','Маленький (1-4)','Средний (4-8)','Большой (8-12)','Огромный (12-16)']
        Button: text: 'Сгенерировать'; on_release: app.generate_current()
    BoxLayout:
        size_hint_y: None; height: dp(42); spacing: dp(5)
        Label: id: merchant_info; text: 'Торговец не создан'; halign: 'left'
        TextInput: id: capital; hint_text: 'Капитал'; input_filter: 'int'; size_hint_x: .35; multiline: False; on_text_validate: app.update_capital()
    Label: text: 'Ассортимент (нажмите на строку для редактирования)'; size_hint_y: None; height: dp(28)
    ScrollView:
        GridLayout: id: items_box; cols: 1; size_hint_y: None; height: self.minimum_height; spacing: dp(3)
    BoxLayout:
        size_hint_y: None; height: dp(42); spacing: dp(5)
        Button: text: 'Выбрать из базы'; on_release: app.open_picker()
        Button: text: 'Удалить выбранное'; on_release: app.remove_selected_item()
    BoxLayout:
        size_hint_y: None; height: dp(42); spacing: dp(4)
        TextInput: id: new_name; hint_text: 'Название'; multiline: False
        TextInput: id: new_qty; hint_text: 'Кол-во'; text: '1'; input_filter: 'int'; multiline: False; size_hint_x: .18
        TextInput: id: new_price; hint_text: 'Цена'; input_filter: 'int'; multiline: False; size_hint_x: .22
        Button: text: 'Добавить'; size_hint_x: .3; on_release: app.add_custom()
    BoxLayout:
        size_hint_y: None; height: dp(44); spacing: dp(5)
        Button: text: 'Сохранить в библиотеку'; on_release: app.save_merchant()
<LibraryScreen>:
    orientation: 'vertical'; padding: dp(8); spacing: dp(5)
    TextInput: id: lib_search; hint_text: 'Поиск торговца'; size_hint_y: None; height: dp(40); multiline: False; on_text: app.refresh_library()
    ScrollView:
        GridLayout: id: library_box; cols: 1; size_hint_y: None; height: self.minimum_height; spacing: dp(4)
    BoxLayout:
        size_hint_y: None; height: dp(44); spacing: dp(4)
        Button: text: 'Загрузить'; on_release: app.load_selected_merchant()
        Button: text: 'Удалить'; on_release: app.delete_selected_merchant()
<DatabaseScreen>:
    orientation: 'vertical'; padding: dp(8); spacing: dp(5)
    TextInput: id: db_search; hint_text: 'Поиск по названию'; size_hint_y: None; height: dp(40); multiline: False; on_text: app.refresh_database()
    BoxLayout:
        size_hint_y: None; height: dp(38)
        Spinner: id: db_sort; text: 'Без сортировки'; values: ['Без сортировки','ID','Название','Цена','Макс. количество','Специализации','Редкость']; on_text: app.refresh_database()
    ScrollView:
        GridLayout: id: db_box; cols: 1; size_hint_y: None; height: self.minimum_height; spacing: dp(2)
'''

class GeneratorScreen(BoxLayout): pass
class LibraryScreen(BoxLayout): pass
class DatabaseScreen(BoxLayout): pass

class MerchantApp(App):
    spec_values = ListProperty(ALL_SPECS)
    def build(self):
        self.title = 'Генератор торговцев One Piece'
        self.root_widget = Builder.load_string(KV)
        self.data_dir = self.user_data_dir
        self.db_path = os.path.join(self.data_dir, 'База_товаров.xlsx')
        bundled = os.path.join(os.path.dirname(__file__), 'assets', 'База_товаров.xlsx')
        if not os.path.exists(self.db_path): shutil.copyfile(bundled, self.db_path)
        self.library_path = os.path.join(self.data_dir, 'торговцы.json')
        self.items = self.load_db(); self.merchant = None; self.selected_index = None; self.selected_library = None
        self.library = self.load_library(); self.refresh_database(); self.refresh_library()
        return self.root_widget
    def load_db(self):
        wb=load_workbook(self.db_path, data_only=True); ws=wb['Товары']; rows=list(ws.iter_rows(values_only=True)); wb.close()
        out=[]
        for r in rows[1:]:
            if not r or not r[0]: continue
            out.append({'id':str(r[0]),'name':str(r[1]),'base_price':int(r[2]),'stack_max':int(r[3]),'tags':[x.strip() for x in str(r[4]).replace(',',';').split(';') if x.strip()],'rarity':str(r[5])})
        return out
    def save_db(self):
        wb=load_workbook(self.db_path); ws=wb['Товары']
        if ws.max_row>1: ws.delete_rows(2,ws.max_row-1)
        for x in self.items: ws.append([x['id'],x['name'],x['base_price'],x['stack_max'],'; '.join(x['tags']),x['rarity']])
        wb.save(self.db_path); wb.close()
    def load_library(self):
        try:
            with open(self.library_path,encoding='utf-8') as f: return json.load(f)
        except Exception: return []
    def save_library_file(self):
        with open(self.library_path,'w',encoding='utf-8') as f: json.dump(self.library,f,ensure_ascii=False,indent=2)
    def popup(self,title,content): Popup(title=title,content=content,size_hint=(.9,.75)).open()
    def generate_current(self):
        s=self.root_widget.ids if hasattr(self.root_widget,'ids') else None
        screen=self.root_widget.tab_list[-1] if False else None
        gen=self.root_widget.tab_list[2] if False else None
        # Access active generator through tab content references
        g=self.generator_screen()
        sr=g.ids.sr.text; seller=g.ids.seller.text; spec=g.ids.spec.text; assort=g.ids.assort.text
        sr=random.choice(RARITIES) if sr=='Случайно' else sr; seller=random.choice(RARITIES) if seller=='Случайно' else seller
        spec=random.choice(SPECS_BY_RARITY[sr]) if spec=='Случайно' else spec
        ranges={'Маленький (1-4)':(1,4),'Средний (4-8)':(4,8),'Большой (8-12)':(8,12),'Огромный (12-16)':(12,16)}
        n=random.randint(*(ranges.get(assort,(1,16))))
        pool=[x for x in self.items if spec in x['tags']]
        if not pool: pool=self.items[:]
        chosen=[]
        for _ in range(min(n,len(pool))):
            it=random.choice([x for x in pool if x not in chosen] or pool)
            rar=random.choices(RARITIES,weights=RARITY_DISTRIBUTION[seller])[0]
            chosen.append({'id':it['id'],'name':it['name'],'qty':random.randint(1,it['stack_max']),'price':max(1,round(it['base_price']*random.uniform(.7,1.3)/10)*10),'rarity':rar})
        lo,hi=CAPITAL_RANGES[seller]
        self.merchant={'id':f'T-{random.randint(100,999)}','name':random.choice(NAMES),'spec':spec,'seller_rarity':seller,'capital':random.randint(lo,hi),'items':chosen}
        g.ids.capital.text=str(self.merchant['capital']); self.refresh_merchant()
    def generator_screen(self): return self.root_widget.tab_list[-1] if False else self._find_screen('GeneratorScreen')
    def _find_screen(self, clsname):
        def walk(w):
            if w.__class__.__name__==clsname: return w
            for c in w.children:
                z=walk(c)
                if z:return z
        return walk(self.root_widget)
    def refresh_merchant(self):
        g=self.generator_screen(); g.ids.merchant_info.text='Торговец не создан' if not self.merchant else f"{self.merchant['name']} · {self.merchant['spec']} · {self.merchant['seller_rarity']}"
        box=g.ids.items_box; box.clear_widgets()
        if not self.merchant:return
        for i,it in enumerate(self.merchant['items']):
            b=Button(text=f"{it['qty']}    {it['name']}    {it['price']:,} белли    {it.get('rarity','Обычный')}".replace(',','.'),size_hint_y=None,height=dp(42),background_color=(.82,.65,.25,1),color=(.16,.1,.04,1))
            b.bind(on_release=lambda _b,idx=i:self.edit_item(idx)); box.add_widget(b)
    def edit_item(self,idx):
        it=self.merchant['items'][idx]; layout=BoxLayout(orientation='vertical',spacing=dp(5),padding=dp(8)); fields=[]
        for key in ('qty','name','price'):
            t=TextInput(text=str(it[key]),multiline=False); layout.add_widget(Label(text={'qty':'Количество','name':'Название','price':'Цена'}[key],size_hint_y=None,height=dp(24))); layout.add_widget(t); fields.append(t)
        btn=Button(text='Сохранить',size_hint_y=None,height=dp(42)); layout.add_widget(btn); p=Popup(title='Редактировать товар',content=layout,size_hint=(.9,.7))
        def save(*_):
            try: it['qty']=int(fields[0].text); it['name']=fields[1].text.strip(); it['price']=int(fields[2].text); p.dismiss(); self.refresh_merchant()
            except: self.popup('Ошибка',Label(text='Проверьте количество и цену'))
        btn.bind(on_release=save); p.open()
    def add_custom(self):
        g=self.generator_screen()
        if not self.merchant:return self.popup('Нет торговца',Label(text='Сначала сгенерируйте торговца'))
        try: self.merchant['items'].append({'id':'CUSTOM-'+str(random.randint(10000,99999)),'name':g.ids.new_name.text.strip(),'qty':int(g.ids.new_qty.text),'price':int(g.ids.new_price.text),'rarity':'Обычный'})
        except: return self.popup('Ошибка',Label(text='Проверьте название, количество и цену'))
        g.ids.new_name.text=''; g.ids.new_qty.text='1'; g.ids.new_price.text=''; self.refresh_merchant()
    def remove_selected_item(self):
        if self.selected_index is not None and self.merchant: self.merchant['items'].pop(self.selected_index); self.selected_index=None; self.refresh_merchant()
        else:self.popup('Удаление',Label(text='Для удаления нажмите и удерживайте товар в списке (или сначала откройте его редактирование).'))
    def open_picker(self):
        if not self.merchant:return self.popup('Нет торговца',Label(text='Сначала сгенерируйте торговца'))
        layout=BoxLayout(orientation='vertical',spacing=dp(5),padding=dp(6)); search=TextInput(hint_text='Поиск по названию',size_hint_y=None,height=dp(40),multiline=False); layout.add_widget(search)
        scroll=ScrollView(); rows=GridLayout(cols=1,size_hint_y=None,spacing=dp(2)); rows.bind(minimum_height=rows.setter('height')); scroll.add_widget(rows); layout.add_widget(scroll)
        checks={}
        def redraw(*_):
            rows.clear_widgets()
            for item in self.items:
                if search.text.lower() not in item['name'].lower():continue
                line=BoxLayout(size_hint_y=None,height=dp(42)); cb=CheckBox(size_hint_x=None,width=dp(45)); checks[item['id']]=cb
                line.add_widget(cb); line.add_widget(Label(text=f"{item['name']} · {item['base_price']} белли",halign='left')); rows.add_widget(line)
        search.bind(text=redraw); redraw(); btn=Button(text='Добавить выбранные',size_hint_y=None,height=dp(44)); layout.add_widget(btn); p=Popup(title='Выбор товаров из базы',content=layout,size_hint=(.96,.9))
        def add(*_):
            for it in self.items:
                if it['id'] in checks and checks[it['id']].active:self.merchant['items'].append({'id':it['id'],'name':it['name'],'qty':1,'price':it['base_price'],'rarity':it['rarity']})
            p.dismiss(); self.refresh_merchant()
        btn.bind(on_release=add); p.open()
    def update_capital(self):
        if self.merchant:
            try:self.merchant['capital']=int(self.generator_screen().ids.capital.text)
            except:pass
    def save_merchant(self):
        if not self.merchant:return self.popup('Нет торговца',Label(text='Сначала создайте торговца'))
        record=dict(self.merchant); record['id']=str(record.get('id') or random.randint(100000,999999)); self.library.append(record); self.save_library_file(); self.refresh_library(); self.popup('Сохранено',Label(text='Торговец добавлен в библиотеку'))
    def refresh_library(self):
        screen=self._find_screen('LibraryScreen'); box=screen.ids.library_box; box.clear_widgets(); q=screen.ids.lib_search.text.lower()
        for idx,m in enumerate(self.library):
            if q and q not in m.get('name','').lower():continue
            b=Button(text=f"{m.get('name')} · {m.get('spec')} · {len(m.get('items',[]))} товаров",size_hint_y=None,height=dp(48))
            b.bind(on_release=lambda _b,i=idx:self.select_library(i)); box.add_widget(b)
    def select_library(self,i): self.selected_library=i
    def load_selected_merchant(self):
        if self.selected_library is None:return
        self.merchant=dict(self.library[self.selected_library]); self.root_widget.switch_to(self.root_widget.tab_list[-1]) if False else None; self.refresh_merchant(); self.popup('Загружено',Label(text='Торговец загружен в генератор'))
    def delete_selected_merchant(self):
        if self.selected_library is None:return
        self.library.pop(self.selected_library); self.selected_library=None; self.save_library_file(); self.refresh_library()
    def refresh_database(self,*_):
        screen=self._find_screen('DatabaseScreen'); box=screen.ids.db_box
        if not hasattr(screen.ids,'db_search'):return
        q=screen.ids.db_search.text.lower(); sort=screen.ids.db_sort.text; data=list(self.items)
        if q:data=[x for x in data if q in x['name'].lower()]
        keys={'ID':'id','Название':'name','Цена':'base_price','Макс. количество':'stack_max','Редкость':'rarity','Специализации':'tags'}
        if sort in keys:data.sort(key=lambda x:str(x[keys[sort]]).lower())
        box.clear_widgets()
        for it in data:
            b=Button(text=f"{it['id']} · {it['name']} · {it['base_price']} · {', '.join(it['tags'])} · {it['rarity']}",size_hint_y=None,height=dp(42))
            b.bind(on_release=lambda _b,x=it:self.edit_db_item(x)); box.add_widget(b)
    def edit_db_item(self,it):
        layout=BoxLayout(orientation='vertical',spacing=dp(4),padding=dp(8)); inputs=[]
        for label,key,val in [('Название','name',it['name']),('Цена','base_price',str(it['base_price'])),('Макс. количество','stack_max',str(it['stack_max'])),('Специализации через ;','tags','; '.join(it['tags']))]:
            layout.add_widget(Label(text=label,size_hint_y=None,height=dp(22))); t=TextInput(text=str(val),multiline=False); layout.add_widget(t); inputs.append(t)
        sp=Spinner(text=it['rarity'],values=RARITIES,size_hint_y=None,height=dp(40)); layout.add_widget(sp); btn=Button(text='Сохранить',size_hint_y=None,height=dp(42)); layout.add_widget(btn); p=Popup(title='Редактировать товар',content=layout,size_hint=(.92,.8))
        def save(*_):
            try:it['name']=inputs[0].text.strip(); it['base_price']=int(inputs[1].text); it['stack_max']=int(inputs[2].text); it['tags']=[x.strip() for x in inputs[3].text.split(';') if x.strip()]; it['rarity']=sp.text; self.save_db(); self.refresh_database(); p.dismiss()
            except:self.popup('Ошибка',Label(text='Проверьте числовые поля'))
        btn.bind(on_release=save); p.open()

if __name__ == '__main__': MerchantApp().run()

/**
 * CINEMA POS - GOOGLE APPS SCRIPT API
 * Database: Google Sheets
 *
 * Deploy this file as a Web App:
 * Execute as: Me
 * Who has access: Anyone
 *
 * Then put the deployed /exec URL into API_URL in Index.html.
 *
 * IMPORTANT:
 * GitHub Pages cannot execute this .gs file. This file is the backend/API
 * that connects the GitHub frontend to Google Sheets.
 *
 * This implementation uses JSONP GET requests so the GitHub static page
 * can communicate cross-origin without requiring a server-side CORS proxy.
 * For a real public commercial deployment, put authentication/rate limiting
 * in front of this API or move the API to a proper backend.
 */

const APP = {
  VERSION: '1.0.0',
  DEFAULT_SHEET_NAME: 'CinemaPOS_DB',
  DEFAULT_ADMIN_USER: 'admin',
  DEFAULT_ADMIN_PASS: 'admin123',
  SESSION_TTL_MS: 8 * 60 * 60 * 1000,
  SHEETS: {
    SETTINGS: 'Settings',
    USERS: 'Users',
    MOVIES: 'Movies',
    HALLS: 'Halls',
    SEATS: 'Seats',
    SHOWTIMES: 'Showtimes',
    SALES: 'Sales',
    SALE_ITEMS: 'SaleItems',
    TICKETS: 'Tickets',
    AUDIT: 'AuditLog'
  }
};

const HEADERS = {
  Settings: ['key','value','updatedAt'],
  Users: ['id','username','password','displayName','role','active','createdAt','updatedAt'],
  Movies: ['id','title','description','durationMin','poster','rating','status','createdAt','updatedAt'],
  Halls: ['id','name','rows','cols','bluePrice','redPrice','screenLabel','status','createdAt','updatedAt'],
  Seats: ['id','hallId','seatCode','rowNo','colNo','seatType','active'],
  Showtimes: ['id','movieId','hallId','showDate','showTime','priceBlue','priceRed','status','createdAt','updatedAt'],
  Sales: ['id','saleNo','showtimeId','movieId','hallId','saleDateTime','staffId','staffName','paymentMethod','subtotal','discount','total','status','createdAt'],
  SaleItems: ['id','saleId','showtimeId','seatId','seatCode','seatType','price','ticketNo','createdAt'],
  Tickets: ['id','ticketNo','saleId','showtimeId','movieId','hallId','seatCode','seatType','movieTitle','showDate','showTime','price','status','createdAt'],
  AuditLog: ['id','userId','username','action','entity','entityId','details','createdAt']
};

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.action === 'health') {
      return jsonResponse_({ok:true, version:APP.VERSION, time:now_()});
    }

    ensureDatabase_();

    const action = String(p.action || 'bootstrap');
    const callback = safeCallback_(p.callback);

    let result;
    switch (action) {
      case 'bootstrap':
        result = apiBootstrap_(p);
        break;
      case 'login':
        result = apiLogin_(p);
        break;
      case 'logout':
        result = apiLogout_(p);
        break;
      case 'movies':
        result = apiList_(APP.SHEETS.MOVIES, p);
        break;
      case 'halls':
        result = apiHalls_(p);
        break;
      case 'showtimes':
        result = apiShowtimes_(p);
        break;
      case 'seatmap':
        result = apiSeatMap_(p);
        break;
      case 'sales':
        result = apiSales_(p);
        break;
      case 'tickets':
        result = apiTickets_(p);
        break;
      case 'dashboard':
        result = apiDashboard_(p);
        break;
      case 'settings':
        result = apiSettings_(p);
        break;
      case 'users':
        result = apiUsers_(p);
        break;
      case 'saveMovie':
        result = apiSaveMovie_(p);
        break;
      case 'deleteMovie':
        result = apiDelete_(APP.SHEETS.MOVIES, p);
        break;
      case 'saveHall':
        result = apiSaveHall_(p);
        break;
      case 'deleteHall':
        result = apiDelete_(APP.SHEETS.HALLS, p);
        break;
      case 'saveShowtime':
        result = apiSaveShowtime_(p);
        break;
      case 'deleteShowtime':
        result = apiDelete_(APP.SHEETS.SHOWTIMES, p);
        break;
      case 'saveUser':
        result = apiSaveUser_(p);
        break;
      case 'deleteUser':
        result = apiDelete_(APP.SHEETS.USERS, p);
        break;
      case 'saveSettings':
        result = apiSaveSettings_(p);
        break;
      case 'createSale':
        result = apiCreateSale_(p);
        break;
      default:
        result = {ok:false,error:'Unknown action: '+action};
    }

    return callback
      ? jsonpResponse_(callback,result)
      : jsonResponse_(result);

  } catch (err) {
    const result = {ok:false,error:String(err && err.message ? err.message : err)};
    const callback = safeCallback_(e && e.parameter && e.parameter.callback);
    return callback ? jsonpResponse_(callback,result) : jsonResponse_(result);
  }
}

function doPost(e) {
  try {
    ensureDatabase_();
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const action = String(body.action || '');
    let result;
    switch (action) {
      case 'createSale':
        result = apiCreateSale_(body);
        break;
      default:
        result = {ok:false,error:'POST action not supported: '+action};
    }
    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({ok:false,error:String(err && err.message ? err.message : err)});
  }
}

function setupDatabase() {
  ensureDatabase_();
  return 'Cinema POS database initialized: ' + APP.DEFAULT_SHEET_NAME;
}

function getDatabaseSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (_) {}
  }

  const ss = SpreadsheetApp.create(APP.DEFAULT_SHEET_NAME);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function ensureDatabase_() {
  const ss = getDatabaseSpreadsheet_();

  Object.keys(HEADERS).forEach(function(name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);

    const headers = HEADERS[name];
    const current = sh.getRange(1,1,1,headers.length).getValues()[0];
    let needs = false;
    for (let i=0;i<headers.length;i++) {
      if (String(current[i] || '') !== headers[i]) {
        needs = true;
        break;
      }
    }
    if (needs) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold');
    }
  });

  const users = sheetRows_(APP.SHEETS.USERS);
  if (!users.length) {
    appendRow_(APP.SHEETS.USERS, {
      id: uid_('USR'),
      username: APP.DEFAULT_ADMIN_USER,
      password: APP.DEFAULT_ADMIN_PASS,
      displayName: 'Administrator',
      role: 'admin',
      active: true,
      createdAt: now_(),
      updatedAt: now_()
    });
  }

  const settings = sheetRows_(APP.SHEETS.SETTINGS);
  const defaults = {
    cinemaName: 'MY CINEMA',
    cinemaAddress: '',
    cinemaPhone: '',
    logo: '',
    ticketHeader: 'ขอบคุณที่ใช้บริการ',
    ticketFooter: 'กรุณาเก็บตั๋วไว้ตลอดการเข้าชม',
    printerWidth: '80',
    paperFeed: '4',
    paperCut: '1'
  };
  Object.keys(defaults).forEach(function(k) {
    if (!settings.some(function(x){ return x.key === k; })) {
      appendRow_(APP.SHEETS.SETTINGS,{key:k,value:defaults[k],updatedAt:now_()});
    }
  });
}

function apiBootstrap_(p) {
  return {
    ok:true,
    version:APP.VERSION,
    settings:settingsObject_(),
    movies:sheetRows_(APP.SHEETS.MOVIES),
    halls:sheetRows_(APP.SHEETS.HALLS),
    showtimes:sheetRows_(APP.SHEETS.SHOWTIMES),
    users: sanitizeUsers_(sheetRows_(APP.SHEETS.USERS)),
    session: p.token ? validateSession_(p.token) : null
  };
}

function apiLogin_(p) {
  const username = String(p.username || '').trim();
  const password = String(p.password || '');
  if (!username || !password) return {ok:false,error:'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน'};

  const user = sheetRows_(APP.SHEETS.USERS).find(function(u){
    return String(u.username) === username && String(u.password) === password && truthy_(u.active);
  });
  if (!user) return {ok:false,error:'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'};

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session_'+token, JSON.stringify({
    id:user.id, username:user.username, displayName:user.displayName, role:user.role
  }), Math.floor(APP.SESSION_TTL_MS/1000));

  audit_(user,'LOGIN','Auth',user.id,'เข้าสู่ระบบ');

  return {
    ok:true,
    token:token,
    user:{id:user.id,username:user.username,displayName:user.displayName,role:user.role}
  };
}

function apiLogout_(p) {
  if (p.token) CacheService.getScriptCache().remove('session_'+p.token);
  return {ok:true};
}

function apiList_(sheet,p) {
  const user = requireSession_(p, false);
  if (!user) return {ok:false,error:'กรุณาเข้าสู่ระบบ'};
  return {ok:true,items:sheetRows_(sheet)};
}

function apiHalls_(p) {
  const user = requireSession_(p,false);
  if (!user) return {ok:false,error:'กรุณาเข้าสู่ระบบ'};
  const halls = sheetRows_(APP.SHEETS.HALLS);
  const seats = sheetRows_(APP.SHEETS.SEATS);
  halls.forEach(function(h){
    h.seats = seats.filter(function(s){return s.hallId === h.id;});
  });
  return {ok:true,items:halls};
}

function apiShowtimes_(p) {
  const user = requireSession_(p,false);
  if (!user) return {ok:false,error:'กรุณาเข้าสู่ระบบ'};
  return {ok:true,items:sheetRows_(APP.SHEETS.SHOWTIMES)};
}

function apiSeatMap_(p) {
  const user = requireSession_(p,false);
  if (!user) return {ok:false,error:'กรุณาเข้าสู่ระบบ'};

  const showtimeId = String(p.showtimeId || '');
  if (!showtimeId) return {ok:false,error:'ไม่พบรอบฉาย'};

  const st = sheetRows_(APP.SHEETS.SHOWTIMES).find(function(x){return x.id===showtimeId;});
  if (!st) return {ok:false,error:'ไม่พบรอบฉาย'};

  const hall = sheetRows_(APP.SHEETS.HALLS).find(function(x){return x.id===st.hallId;});
  if (!hall) return {ok:false,error:'ไม่พบโรงภาพยนตร์'};

  const seats = sheetRows_(APP.SHEETS.SEATS).filter(function(x){return x.hallId===hall.id && truthy_(x.active);});
  const sold = sheetRows_(APP.SHEETS.SALE_ITEMS)
    .filter(function(x){return x.showtimeId===showtimeId;})
    .map(function(x){return x.seatId;});

  const soldSet = {};
  sold.forEach(function(x){soldSet[x]=true;});

  seats.forEach(function(s){s.sold=!!soldSet[s.id];});
  return {ok:true,hall:hall,seats:seats};
}

function apiSales_(p) {
  const user = requireSession_(p,false);
  if (!user) return {ok:false,error:'กรุณาเข้าสู่ระบบ'};
  return {ok:true,items:sheetRows_(APP.SHEETS.SALES)};
}

function apiTickets_(p) {
  const user = requireSession_(p,false);
  if (!user) return {ok:false,error:'กรุณาเข้าสู่ระบบ'};
  return {ok:true,items:sheetRows_(APP.SHEETS.TICKETS)};
}

function apiDashboard_(p) {
  const user = requireSession_(p,false);
  if (!user) return {ok:false,error:'กรุณาเข้าสู่ระบบ'};

  const sales = sheetRows_(APP.SHEETS.SALES);
  const tickets = sheetRows_(APP.SHEETS.TICKETS);
  const movies = sheetRows_(APP.SHEETS.MOVIES);
  const showtimes = sheetRows_(APP.SHEETS.SHOWTIMES);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const todaySales = sales.filter(function(x){return String(x.saleDateTime).indexOf(today)===0 && x.status!=='CANCELLED';});
  const revenue = todaySales.reduce(function(a,x){return a+num_(x.total);},0);

  return {
    ok:true,
    cards:{
      movies:movies.filter(function(x){return x.status!=='INACTIVE';}).length,
      showtimes:showtimes.filter(function(x){return x.status!=='CANCELLED';}).length,
      todaySales:todaySales.length,
      todayRevenue:revenue,
      tickets: tickets.filter(function(x){return x.status!=='CANCELLED';}).length
    }
  };
}

function apiSettings_(p) {
  const user = requireSession_(p,false);
  if (!user) return {ok:false,error:'กรุณาเข้าสู่ระบบ'};
  return {ok:true,settings:settingsObject_()};
}

function apiUsers_(p) {
  const user = requireSession_(p,true);
  if (!user) return {ok:false,error:'ไม่มีสิทธิ์'};
  return {ok:true,items:sanitizeUsers_(sheetRows_(APP.SHEETS.USERS))};
}

function apiSaveMovie_(p) {
  const user = requireSession_(p,true);
  if (!user) return {ok:false,error:'ไม่มีสิทธิ์'};
  const item = parseJsonParam_(p.data);
  if (!item.title) return {ok:false,error:'กรุณาระบุชื่อหนัง'};
  item.id = item.id || uid_('MOV');
  item.status = item.status || 'ACTIVE';
  upsert_(APP.SHEETS.MOVIES,item,HEADERS.Movies);
  audit_(user,'SAVE','Movie',item.id,'บันทึกภาพยนตร์ '+item.title);
  return {ok:true,item:item};
}

function apiSaveHall_(p) {
  const user = requireSession_(p,true);
  if (!user) return {ok:false,error:'ไม่มีสิทธิ์'};
  const item = parseJsonParam_(p.data);
  const rows = Math.max(1,parseInt(item.rows||8,10));
  const cols = Math.max(1,parseInt(item.cols||12,10));
  item.id = item.id || uid_('HAL');
  item.rows=rows; item.cols=cols;
  item.status=item.status||'ACTIVE';
  item.bluePrice=num_(item.bluePrice);
  item.redPrice=num_(item.redPrice);
  upsert_(APP.SHEETS.HALLS,item,HEADERS.Halls);
  generateSeats_(item);
  audit_(user,'SAVE','Hall',item.id,'บันทึกโรง '+item.name);
  return {ok:true,item:item};
}

function apiSaveShowtime_(p) {
  const user = requireSession_(p,true);
  if (!user) return {ok:false,error:'ไม่มีสิทธิ์'};
  const item = parseJsonParam_(p.data);
  if (!item.movieId || !item.hallId || !item.showDate || !item.showTime) {
    return {ok:false,error:'ข้อมูลรอบฉายไม่ครบ'};
  }
  item.id=item.id||uid_('ST');
  item.priceBlue=num_(item.priceBlue);
  item.priceRed=num_(item.priceRed);
  item.status=item.status||'ACTIVE';
  upsert_(APP.SHEETS.SHOWTIMES,item,HEADERS.Showtimes);
  audit_(user,'SAVE','Showtime',item.id,'บันทึกรอบฉาย');
  return {ok:true,item:item};
}

function apiSaveUser_(p) {
  const user = requireSession_(p,true);
  if (!user) return {ok:false,error:'ไม่มีสิทธิ์'};
  const item = parseJsonParam_(p.data);
  if (!item.username || !item.password) return {ok:false,error:'กรุณากรอกข้อมูลผู้ใช้'};
  item.id=item.id||uid_('USR');
  item.active = item.active !== false;
  item.updatedAt=now_();
  if (!item.createdAt) item.createdAt=now_();
  upsert_(APP.SHEETS.USERS,item,HEADERS.Users);
  audit_(user,'SAVE','User',item.id,'บันทึกผู้ใช้งาน');
  return {ok:true,item:sanitizeUser_(item)};
}

function apiSaveSettings_(p) {
  const user = requireSession_(p,true);
  if (!user) return {ok:false,error:'ไม่มีสิทธิ์'};
  const data=parseJsonParam_(p.data);
  const ss=getDatabaseSpreadsheet_();
  const sh=ss.getSheetByName(APP.SHEETS.SETTINGS);
  const rows=sheetRows_(APP.SHEETS.SETTINGS);
  Object.keys(data).forEach(function(key){
    const idx=rows.findIndex(function(x){return x.key===key;});
    if (idx>=0) {
      sh.getRange(idx+2,2).setValue(data[key]);
      sh.getRange(idx+2,3).setValue(now_());
    } else {
      appendRow_(APP.SHEETS.SETTINGS,{key:key,value:data[key],updatedAt:now_()});
    }
  });
  audit_(user,'SAVE','Settings','SETTINGS','แก้ไขการตั้งค่า');
  return {ok:true,settings:settingsObject_()};
}

function apiDelete_(sheet,p) {
  const user=requireSession_(p,true);
  if (!user) return {ok:false,error:'ไม่มีสิทธิ์'};
  const id=String(p.id||'');
  if (!id) return {ok:false,error:'ไม่พบรหัสรายการ'};
  const sh=getDatabaseSpreadsheet_().getSheetByName(sheet);
  const rows=sheetRows_(sheet);
  const idx=rows.findIndex(function(x){return x.id===id;});
  if (idx<0) return {ok:false,error:'ไม่พบรายการ'};
  sh.deleteRow(idx+2);
  audit_(user,'DELETE',sheet,id,'ลบรายการ');
  return {ok:true};
}

function apiCreateSale_(p) {
  const user=requireSession_(p,true);
  if (!user) return {ok:false,error:'ไม่มีสิทธิ์'};

  const data=parseJsonParam_(p.data);
  const showtimeId=String(data.showtimeId||'');
  const selectedSeats=Array.isArray(data.seats)?data.seats:[];
  const paymentMethod=String(data.paymentMethod||'CASH');

  if (!showtimeId || !selectedSeats.length) return {ok:false,error:'กรุณาเลือกที่นั่ง'};

  const lock=LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const showtimes=sheetRows_(APP.SHEETS.SHOWTIMES);
    const st=showtimes.find(function(x){return x.id===showtimeId;});
    if (!st) return {ok:false,error:'ไม่พบรอบฉาย'};

    const movie=sheetRows_(APP.SHEETS.MOVIES).find(function(x){return x.id===st.movieId;});
    const hall=sheetRows_(APP.SHEETS.HALLS).find(function(x){return x.id===st.hallId;});
    if (!movie || !hall) return {ok:false,error:'ข้อมูลหนังหรือโรงไม่ครบ'};

    const seats=sheetRows_(APP.SHEETS.SEATS).filter(function(x){return x.hallId===hall.id;});
    const existing=sheetRows_(APP.SHEETS.SALE_ITEMS).filter(function(x){return x.showtimeId===showtimeId;});
    const sold={};
    existing.forEach(function(x){sold[x.seatId]=true;});

    const normalized=[];
    for (let i=0;i<selectedSeats.length;i++) {
      const requested=selectedSeats[i];
      const seatId=String(requested.seatId||requested.id||'');
      const seat=seats.find(function(x){return x.id===seatId;});
      if (!seat) return {ok:false,error:'ไม่พบที่นั่ง '+seatId};
      if (sold[seat.id]) return {ok:false,error:'ที่นั่ง '+seat.seatCode+' ถูกขายไปแล้ว กรุณาโหลดผังใหม่'};
      normalized.push(seat);
      sold[seat.id]=true;
    }

    let subtotal=0;
    normalized.forEach(function(s){
      subtotal += String(s.seatType).toUpperCase()==='RED' ? num_(st.priceRed) : num_(st.priceBlue);
    });

    const discount=num_(data.discount);
    const total=Math.max(0,subtotal-discount);
    const saleId=uid_('SAL');
    const saleNo=makeNo_('S');
    const created=now_();

    appendRow_(APP.SHEETS.SALES,{
      id:saleId,
      saleNo:saleNo,
      showtimeId:st.id,
      movieId:movie.id,
      hallId:hall.id,
      saleDateTime:created,
      staffId:user.id,
      staffName:user.displayName,
      paymentMethod:paymentMethod,
      subtotal:subtotal,
      discount:discount,
      total:total,
      status:'PAID',
      createdAt:created
    });

    const ticketObjects=[];
    normalized.forEach(function(seat){
      const price=String(seat.seatType).toUpperCase()==='RED' ? num_(st.priceRed) : num_(st.priceBlue);
      const ticketNo=makeNo_('T');
      appendRow_(APP.SHEETS.SALE_ITEMS,{
        id:uid_('ITM'),
        saleId:saleId,
        showtimeId:st.id,
        seatId:seat.id,
        seatCode:seat.seatCode,
        seatType:seat.seatType,
        price:price,
        ticketNo:ticketNo,
        createdAt:created
      });
      appendRow_(APP.SHEETS.TICKETS,{
        id:uid_('TKT'),
        ticketNo:ticketNo,
        saleId:saleId,
        showtimeId:st.id,
        movieId:movie.id,
        hallId:hall.id,
        seatCode:seat.seatCode,
        seatType:seat.seatType,
        movieTitle:movie.title,
        showDate:st.showDate,
        showTime:st.showTime,
        price:price,
        status:'VALID',
        createdAt:created
      });
      ticketObjects.push({
        ticketNo:ticketNo,
        seatCode:seat.seatCode,
        seatType:seat.seatType,
        price:price
      });
    });

    audit_(user,'CREATE','Sale',saleId,'ขายตั๋ว '+saleNo);
    return {
      ok:true,
      sale:{
        id:saleId,saleNo:saleNo,
        movieTitle:movie.title,
        showDate:st.showDate,
        showTime:st.showTime,
        hallName:hall.name,
        paymentMethod:paymentMethod,
        subtotal:subtotal,
        discount:discount,
        total:total,
        tickets:ticketObjects
      },
      settings:settingsObject_()
    };
  } finally {
    lock.releaseLock();
  }
}

function generateSeats_(hall) {
  const existing=sheetRows_(APP.SHEETS.SEATS).filter(function(x){return x.hallId===hall.id;});
  const existingCodes={};
  existing.forEach(function(x){existingCodes[x.seatCode]=true;});

  for (let r=1;r<=hall.rows;r++) {
    for (let c=1;c<=hall.cols;c++) {
      const code=numberToLetters_(r)+c;
      if (!existingCodes[code]) {
        const middle=Math.ceil(hall.cols/2);
        const type=c<=middle?'BLUE':'RED';
        appendRow_(APP.SHEETS.SEATS,{
          id:uid_('SET'),
          hallId:hall.id,
          seatCode:code,
          rowNo:r,
          colNo:c,
          seatType:type,
          active:true
        });
      }
    }
  }
}

function settingsObject_() {
  const obj={};
  sheetRows_(APP.SHEETS.SETTINGS).forEach(function(x){obj[x.key]=x.value;});
  return obj;
}

function sanitizeUsers_(users) {
  return users.map(sanitizeUser_);
}

function sanitizeUser_(u) {
  return {
    id:u.id,username:u.username,displayName:u.displayName,
    role:u.role,active:truthy_(u.active),createdAt:u.createdAt,updatedAt:u.updatedAt
  };
}

function validateSession_(token) {
  if (!token) return null;
  const raw=CacheService.getScriptCache().get('session_'+token);
  if (!raw) return null;
  return JSON.parse(raw);
}

function requireSession_(p,adminOnly) {
  const s=validateSession_(String((p||{}).token||''));
  if (!s) return null;
  if (adminOnly && s.role!=='admin') return null;
  return s;
}

function audit_(user,action,entity,entityId,details) {
  appendRow_(APP.SHEETS.AUDIT_LOG,{
    id:uid_('LOG'),
    userId:user.id,
    username:user.username,
    action:action,
    entity:entity,
    entityId:entityId,
    details:details,
    createdAt:now_()
  });
}

function sheetRows_(name) {
  const sh=getDatabaseSpreadsheet_().getSheetByName(name);
  if (!sh || sh.getLastRow()<2) return [];
  const values=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getValues();
  const headers=values[0];
  return values.slice(1).filter(function(row){
    return row.some(function(v){return v!=='' && v!==null;});
  }).map(function(row){
    const obj={};
    headers.forEach(function(h,i){
      let v=row[i];
      if (v instanceof Date) v=Utilities.formatDate(v,Session.getScriptTimeZone(),"yyyy-MM-dd HH:mm:ss");
      obj[h]=v;
    });
    return obj;
  });
}

function appendRow_(sheet,obj) {
  const headers=HEADERS[sheet];
  const sh=getDatabaseSpreadsheet_().getSheetByName(sheet);
  sh.appendRow(headers.map(function(h){
    return obj[h] === undefined ? '' : obj[h];
  }));
}

function upsert_(sheet,obj,headers) {
  const sh=getDatabaseSpreadsheet_().getSheetByName(sheet);
  const rows=sheetRows_(sheet);
  const key=obj.id;
  const idx=rows.findIndex(function(x){return x.id===key;});
  if (idx<0) {
    appendRow_(sheet,obj);
  } else {
    sh.getRange(idx+2,1,1,headers.length).setValues([headers.map(function(h){
      return obj[h] === undefined ? rows[idx][h] || '' : obj[h];
    })]);
  }
}

function parseJsonParam_(v) {
  if (!v) return {};
  if (typeof v === 'object') return v;
  try {return JSON.parse(v);} catch (_) {throw new Error('รูปแบบข้อมูลไม่ถูกต้อง');}
}

function uid_(prefix) {
  return prefix+'_'+Utilities.getUuid().replace(/-/g,'').slice(0,18).toUpperCase();
}

function makeNo_(prefix) {
  const d=new Date();
  const date=Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyyMMdd');
  const rand=Math.floor(100000+Math.random()*900000);
  return prefix+date+rand;
}

function now_() {
  return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
}

function num_(v) {
  const n=Number(v);
  return isFinite(n)?n:0;
}

function truthy_(v) {
  return v===true || v===1 || String(v).toLowerCase()==='true';
}

function numberToLetters_(n) {
  let s='';
  while(n>0) {
    const r=(n-1)%26;
    s=String.fromCharCode(65+r)+s;
    n=Math.floor((n-1)/26);
  }
  return s;
}

function safeCallback_(v) {
  const c=String(v||'');
  return /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(c) ? c : '';
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpResponse_(callback,obj) {
  return ContentService.createTextOutput(callback+'('+JSON.stringify(obj)+')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

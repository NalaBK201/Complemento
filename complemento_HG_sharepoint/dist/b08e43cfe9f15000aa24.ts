function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t.return || t.return(); } finally { if (u) throw o; } } }; }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _regeneratorRuntime() { "use strict"; /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/facebook/regenerator/blob/main/LICENSE */ _regeneratorRuntime = function _regeneratorRuntime() { return e; }; var t, e = {}, r = Object.prototype, n = r.hasOwnProperty, o = Object.defineProperty || function (t, e, r) { t[e] = r.value; }, i = "function" == typeof Symbol ? Symbol : {}, a = i.iterator || "@@iterator", c = i.asyncIterator || "@@asyncIterator", u = i.toStringTag || "@@toStringTag"; function define(t, e, r) { return Object.defineProperty(t, e, { value: r, enumerable: !0, configurable: !0, writable: !0 }), t[e]; } try { define({}, ""); } catch (t) { define = function define(t, e, r) { return t[e] = r; }; } function wrap(t, e, r, n) { var i = e && e.prototype instanceof Generator ? e : Generator, a = Object.create(i.prototype), c = new Context(n || []); return o(a, "_invoke", { value: makeInvokeMethod(t, r, c) }), a; } function tryCatch(t, e, r) { try { return { type: "normal", arg: t.call(e, r) }; } catch (t) { return { type: "throw", arg: t }; } } e.wrap = wrap; var h = "suspendedStart", l = "suspendedYield", f = "executing", s = "completed", y = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} var p = {}; define(p, a, function () { return this; }); var d = Object.getPrototypeOf, v = d && d(d(values([]))); v && v !== r && n.call(v, a) && (p = v); var g = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(p); function defineIteratorMethods(t) { ["next", "throw", "return"].forEach(function (e) { define(t, e, function (t) { return this._invoke(e, t); }); }); } function AsyncIterator(t, e) { function invoke(r, o, i, a) { var c = tryCatch(t[r], t, o); if ("throw" !== c.type) { var u = c.arg, h = u.value; return h && "object" == _typeof(h) && n.call(h, "__await") ? e.resolve(h.__await).then(function (t) { invoke("next", t, i, a); }, function (t) { invoke("throw", t, i, a); }) : e.resolve(h).then(function (t) { u.value = t, i(u); }, function (t) { return invoke("throw", t, i, a); }); } a(c.arg); } var r; o(this, "_invoke", { value: function value(t, n) { function callInvokeWithMethodAndArg() { return new e(function (e, r) { invoke(t, n, e, r); }); } return r = r ? r.then(callInvokeWithMethodAndArg, callInvokeWithMethodAndArg) : callInvokeWithMethodAndArg(); } }); } function makeInvokeMethod(e, r, n) { var o = h; return function (i, a) { if (o === f) throw Error("Generator is already running"); if (o === s) { if ("throw" === i) throw a; return { value: t, done: !0 }; } for (n.method = i, n.arg = a;;) { var c = n.delegate; if (c) { var u = maybeInvokeDelegate(c, n); if (u) { if (u === y) continue; return u; } } if ("next" === n.method) n.sent = n._sent = n.arg;else if ("throw" === n.method) { if (o === h) throw o = s, n.arg; n.dispatchException(n.arg); } else "return" === n.method && n.abrupt("return", n.arg); o = f; var p = tryCatch(e, r, n); if ("normal" === p.type) { if (o = n.done ? s : l, p.arg === y) continue; return { value: p.arg, done: n.done }; } "throw" === p.type && (o = s, n.method = "throw", n.arg = p.arg); } }; } function maybeInvokeDelegate(e, r) { var n = r.method, o = e.iterator[n]; if (o === t) return r.delegate = null, "throw" === n && e.iterator.return && (r.method = "return", r.arg = t, maybeInvokeDelegate(e, r), "throw" === r.method) || "return" !== n && (r.method = "throw", r.arg = new TypeError("The iterator does not provide a '" + n + "' method")), y; var i = tryCatch(o, e.iterator, r.arg); if ("throw" === i.type) return r.method = "throw", r.arg = i.arg, r.delegate = null, y; var a = i.arg; return a ? a.done ? (r[e.resultName] = a.value, r.next = e.nextLoc, "return" !== r.method && (r.method = "next", r.arg = t), r.delegate = null, y) : a : (r.method = "throw", r.arg = new TypeError("iterator result is not an object"), r.delegate = null, y); } function pushTryEntry(t) { var e = { tryLoc: t[0] }; 1 in t && (e.catchLoc = t[1]), 2 in t && (e.finallyLoc = t[2], e.afterLoc = t[3]), this.tryEntries.push(e); } function resetTryEntry(t) { var e = t.completion || {}; e.type = "normal", delete e.arg, t.completion = e; } function Context(t) { this.tryEntries = [{ tryLoc: "root" }], t.forEach(pushTryEntry, this), this.reset(!0); } function values(e) { if (e || "" === e) { var r = e[a]; if (r) return r.call(e); if ("function" == typeof e.next) return e; if (!isNaN(e.length)) { var o = -1, i = function next() { for (; ++o < e.length;) if (n.call(e, o)) return next.value = e[o], next.done = !1, next; return next.value = t, next.done = !0, next; }; return i.next = i; } } throw new TypeError(_typeof(e) + " is not iterable"); } return GeneratorFunction.prototype = GeneratorFunctionPrototype, o(g, "constructor", { value: GeneratorFunctionPrototype, configurable: !0 }), o(GeneratorFunctionPrototype, "constructor", { value: GeneratorFunction, configurable: !0 }), GeneratorFunction.displayName = define(GeneratorFunctionPrototype, u, "GeneratorFunction"), e.isGeneratorFunction = function (t) { var e = "function" == typeof t && t.constructor; return !!e && (e === GeneratorFunction || "GeneratorFunction" === (e.displayName || e.name)); }, e.mark = function (t) { return Object.setPrototypeOf ? Object.setPrototypeOf(t, GeneratorFunctionPrototype) : (t.__proto__ = GeneratorFunctionPrototype, define(t, u, "GeneratorFunction")), t.prototype = Object.create(g), t; }, e.awrap = function (t) { return { __await: t }; }, defineIteratorMethods(AsyncIterator.prototype), define(AsyncIterator.prototype, c, function () { return this; }), e.AsyncIterator = AsyncIterator, e.async = function (t, r, n, o, i) { void 0 === i && (i = Promise); var a = new AsyncIterator(wrap(t, r, n, o), i); return e.isGeneratorFunction(r) ? a : a.next().then(function (t) { return t.done ? t.value : a.next(); }); }, defineIteratorMethods(g), define(g, u, "Generator"), define(g, a, function () { return this; }), define(g, "toString", function () { return "[object Generator]"; }), e.keys = function (t) { var e = Object(t), r = []; for (var n in e) r.push(n); return r.reverse(), function next() { for (; r.length;) { var t = r.pop(); if (t in e) return next.value = t, next.done = !1, next; } return next.done = !0, next; }; }, e.values = values, Context.prototype = { constructor: Context, reset: function reset(e) { if (this.prev = 0, this.next = 0, this.sent = this._sent = t, this.done = !1, this.delegate = null, this.method = "next", this.arg = t, this.tryEntries.forEach(resetTryEntry), !e) for (var r in this) "t" === r.charAt(0) && n.call(this, r) && !isNaN(+r.slice(1)) && (this[r] = t); }, stop: function stop() { this.done = !0; var t = this.tryEntries[0].completion; if ("throw" === t.type) throw t.arg; return this.rval; }, dispatchException: function dispatchException(e) { if (this.done) throw e; var r = this; function handle(n, o) { return a.type = "throw", a.arg = e, r.next = n, o && (r.method = "next", r.arg = t), !!o; } for (var o = this.tryEntries.length - 1; o >= 0; --o) { var i = this.tryEntries[o], a = i.completion; if ("root" === i.tryLoc) return handle("end"); if (i.tryLoc <= this.prev) { var c = n.call(i, "catchLoc"), u = n.call(i, "finallyLoc"); if (c && u) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } else if (c) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); } else { if (!u) throw Error("try statement without catch or finally"); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } } } }, abrupt: function abrupt(t, e) { for (var r = this.tryEntries.length - 1; r >= 0; --r) { var o = this.tryEntries[r]; if (o.tryLoc <= this.prev && n.call(o, "finallyLoc") && this.prev < o.finallyLoc) { var i = o; break; } } i && ("break" === t || "continue" === t) && i.tryLoc <= e && e <= i.finallyLoc && (i = null); var a = i ? i.completion : {}; return a.type = t, a.arg = e, i ? (this.method = "next", this.next = i.finallyLoc, y) : this.complete(a); }, complete: function complete(t, e) { if ("throw" === t.type) throw t.arg; return "break" === t.type || "continue" === t.type ? this.next = t.arg : "return" === t.type ? (this.rval = this.arg = t.arg, this.method = "return", this.next = "end") : "normal" === t.type && e && (this.next = e), y; }, finish: function finish(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.finallyLoc === t) return this.complete(r.completion, r.afterLoc), resetTryEntry(r), y; } }, catch: function _catch(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.tryLoc === t) { var n = r.completion; if ("throw" === n.type) { var o = n.arg; resetTryEntry(r); } return o; } } throw Error("illegal catch attempt"); }, delegateYield: function delegateYield(e, r, n) { return this.delegate = { iterator: values(e), resultName: r, nextLoc: n }, "next" === this.method && (this.arg = t), y; } }, e; }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
import Sortable from "sortablejs";
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var _document$getElementB, _document$getElementB2, _document$getElementB3, _document$getElementB4;
    mostrarEncabezadosComoBotones("nombre_tabla");
    cargarOpcionesDesdeEncabezados("T_Filtros", "filtrosPredeterminados");
    (_document$getElementB = document.getElementById("visualizarTodo")) === null || _document$getElementB === void 0 || _document$getElementB.addEventListener("click", function () {
      cambiarVisibilidadColumnasDesdeCelda("nombre_tabla", false); // mostrar columnas
    });
    (_document$getElementB2 = document.getElementById("ocultarTodo")) === null || _document$getElementB2 === void 0 || _document$getElementB2.addEventListener("click", function () {
      cambiarVisibilidadColumnasDesdeCelda("nombre_tabla", true); // ocultar columnas
    });
    (_document$getElementB3 = document.getElementById("filtrosPredeterminados")) === null || _document$getElementB3 === void 0 || _document$getElementB3.addEventListener("change", function () {
      aplicarVistaDesdeTFiltros("nombre_tabla", "filtrosPredeterminados"); // usa el nombre real de tu tabla principal
    });
    (_document$getElementB4 = document.getElementById("guardar")) === null || _document$getElementB4 === void 0 || _document$getElementB4.addEventListener("click", function () {
      guardarVistaSeleccionada("filtrosPredeterminados", "buttonsContainer", "T_Filtros");
    });
  }
});
function mostrarEncabezadosComoBotones(_x) {
  return _mostrarEncabezadosComoBotones.apply(this, arguments);
}
function _mostrarEncabezadosComoBotones() {
  _mostrarEncabezadosComoBotones = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee4(nombreCelda) {
    return _regeneratorRuntime().wrap(function _callee4$(_context5) {
      while (1) switch (_context5.prev = _context5.next) {
        case 0:
          _context5.next = 2;
          return Excel.run(/*#__PURE__*/function () {
            var _ref = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee3(context) {
              var workbook, namedItem, rango, nombreTabla, tabla, encabezadosRange, encabezados, columnas, container, _loop, i;
              return _regeneratorRuntime().wrap(function _callee3$(_context4) {
                while (1) switch (_context4.prev = _context4.next) {
                  case 0:
                    workbook = context.workbook; // 1. Obtener el nombre de la tabla desde la celda nombrada
                    namedItem = workbook.names.getItem(nombreCelda);
                    rango = namedItem.getRange();
                    rango.load("values");
                    _context4.next = 6;
                    return context.sync();
                  case 6:
                    nombreTabla = rango.values[0][0]; // 2. Obtener la tabla y sus encabezados
                    tabla = workbook.tables.getItem(nombreTabla);
                    encabezadosRange = tabla.getHeaderRowRange();
                    encabezadosRange.load("values");
                    tabla.columns.load("items");
                    _context4.next = 13;
                    return context.sync();
                  case 13:
                    encabezados = encabezadosRange.values[0];
                    columnas = tabla.columns.items;
                    container = document.getElementById("buttonsContainer");
                    container.innerHTML = "";
                    _loop = /*#__PURE__*/_regeneratorRuntime().mark(function _loop(i) {
                      var encabezado, columna, rango, boton;
                      return _regeneratorRuntime().wrap(function _loop$(_context3) {
                        while (1) switch (_context3.prev = _context3.next) {
                          case 0:
                            encabezado = encabezados[i];
                            columna = columnas[i];
                            rango = columna.getRange();
                            rango.load("columnHidden");
                            _context3.next = 6;
                            return context.sync();
                          case 6:
                            boton = document.createElement("button");
                            boton.textContent = encabezado;

                            // Estilo base
                            boton.style.textAlign = "left";
                            boton.style.width = "100%";
                            boton.style.padding = "2px 5px";
                            boton.style.margin = "1px 0";
                            boton.style.minWidth = "5px";
                            boton.style.border = "1px solid #aaa";
                            boton.style.borderRadius = "4px";
                            boton.style.cursor = "pointer";
                            boton.style.fontSize = "12px";
                            boton.style.backgroundColor = rango.columnHidden ? "#f5b5b5" : "#cce5cc";
                            boton.style.color = "black";
                            boton.onclick = /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee2() {
                              return _regeneratorRuntime().wrap(function _callee2$(_context2) {
                                while (1) switch (_context2.prev = _context2.next) {
                                  case 0:
                                    _context2.next = 2;
                                    return Excel.run(/*#__PURE__*/function () {
                                      var _ref3 = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee(ctx) {
                                        var col, rangoCol, oculta;
                                        return _regeneratorRuntime().wrap(function _callee$(_context) {
                                          while (1) switch (_context.prev = _context.next) {
                                            case 0:
                                              col = ctx.workbook.tables.getItem(nombreTabla).columns.getItemAt(i);
                                              rangoCol = col.getRange();
                                              rangoCol.load("columnHidden");
                                              _context.next = 5;
                                              return ctx.sync();
                                            case 5:
                                              oculta = rangoCol.columnHidden;
                                              rangoCol.columnHidden = !oculta;
                                              _context.next = 9;
                                              return ctx.sync();
                                            case 9:
                                              // Actualizar estilo botón
                                              boton.style.backgroundColor = !oculta ? "#f5b5b5" : "#cce5cc";
                                            case 10:
                                            case "end":
                                              return _context.stop();
                                          }
                                        }, _callee);
                                      }));
                                      return function (_x14) {
                                        return _ref3.apply(this, arguments);
                                      };
                                    }());
                                  case 2:
                                  case "end":
                                    return _context2.stop();
                                }
                              }, _callee2);
                            }));
                            container.appendChild(boton);
                          case 21:
                          case "end":
                            return _context3.stop();
                        }
                      }, _loop);
                    });
                    i = 0;
                  case 19:
                    if (!(i < columnas.length)) {
                      _context4.next = 24;
                      break;
                    }
                    return _context4.delegateYield(_loop(i), "t0", 21);
                  case 21:
                    i++;
                    _context4.next = 19;
                    break;
                  case 24:
                    // Activar drag & drop con SortableJS
                    Sortable.create(container, {
                      animation: 150,
                      onEnd: function onEnd() {
                        var nuevoOrden = Array.from(container.children).map(function (el) {
                          var _el$textContent;
                          return ((_el$textContent = el.textContent) === null || _el$textContent === void 0 ? void 0 : _el$textContent.trim()) || "";
                        });
                        reordenarColumnasEnExcel(nombreTabla, nuevoOrden);
                      }
                    });
                  case 25:
                  case "end":
                    return _context4.stop();
                }
              }, _callee3);
            }));
            return function (_x13) {
              return _ref.apply(this, arguments);
            };
          }());
        case 2:
        case "end":
          return _context5.stop();
      }
    }, _callee4);
  }));
  return _mostrarEncabezadosComoBotones.apply(this, arguments);
}
function cambiarVisibilidadColumnasDesdeCelda(_x2, _x3) {
  return _cambiarVisibilidadColumnasDesdeCelda.apply(this, arguments);
}
function _cambiarVisibilidadColumnasDesdeCelda() {
  _cambiarVisibilidadColumnasDesdeCelda = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee6(nombreCelda, ocultar) {
    return _regeneratorRuntime().wrap(function _callee6$(_context7) {
      while (1) switch (_context7.prev = _context7.next) {
        case 0:
          _context7.next = 2;
          return Excel.run(/*#__PURE__*/function () {
            var _ref4 = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee5(context) {
              var workbook, namedItem, rango, nombreTabla, hoja, tabla, columnas, _iterator, _step, columna, _rango;
              return _regeneratorRuntime().wrap(function _callee5$(_context6) {
                while (1) switch (_context6.prev = _context6.next) {
                  case 0:
                    workbook = context.workbook; // 1. Obtener el nombre de la tabla desde la celda nombrada
                    namedItem = workbook.names.getItem(nombreCelda);
                    rango = namedItem.getRange();
                    rango.load("values");
                    _context6.next = 6;
                    return context.sync();
                  case 6:
                    nombreTabla = rango.values[0][0]; // 2. Obtener columnas de la tabla
                    hoja = context.workbook.worksheets.getActiveWorksheet();
                    tabla = hoja.tables.getItem(nombreTabla);
                    columnas = tabla.columns;
                    columnas.load("items");
                    _context6.next = 13;
                    return context.sync();
                  case 13:
                    _iterator = _createForOfIteratorHelper(columnas.items);
                    try {
                      for (_iterator.s(); !(_step = _iterator.n()).done;) {
                        columna = _step.value;
                        _rango = columna.getRange();
                        _rango.columnHidden = ocultar;
                      }
                    } catch (err) {
                      _iterator.e(err);
                    } finally {
                      _iterator.f();
                    }
                    _context6.next = 17;
                    return context.sync();
                  case 17:
                    // 3. Pintar botones según visibilidad
                    if (ocultar) {
                      pintarBotonesEnRojo();
                    } else {
                      pintarBotonesEnVerde();
                    }
                  case 18:
                  case "end":
                    return _context6.stop();
                }
              }, _callee5);
            }));
            return function (_x15) {
              return _ref4.apply(this, arguments);
            };
          }());
        case 2:
        case "end":
          return _context7.stop();
      }
    }, _callee6);
  }));
  return _cambiarVisibilidadColumnasDesdeCelda.apply(this, arguments);
}
function pintarBotonesEnVerde() {
  var botones = document.querySelectorAll("#buttonsContainer button");
  botones.forEach(function (boton) {
    boton.style.backgroundColor = "#cce5cc";
  });
}
function pintarBotonesEnRojo() {
  var botones = document.querySelectorAll("#buttonsContainer button");
  botones.forEach(function (boton) {
    boton.style.backgroundColor = "#f5b5b5";
  });
}
function cargarOpcionesDesdeEncabezados(_x4, _x5) {
  return _cargarOpcionesDesdeEncabezados.apply(this, arguments);
}
function _cargarOpcionesDesdeEncabezados() {
  _cargarOpcionesDesdeEncabezados = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee8(nombreTabla, idSelect) {
    return _regeneratorRuntime().wrap(function _callee8$(_context9) {
      while (1) switch (_context9.prev = _context9.next) {
        case 0:
          _context9.next = 2;
          return Excel.run(/*#__PURE__*/function () {
            var _ref5 = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee7(context) {
              var hoja, tabla, encabezadosRange, encabezados, select, boton, opcionDefecto, _boton, _select;
              return _regeneratorRuntime().wrap(function _callee7$(_context8) {
                while (1) switch (_context8.prev = _context8.next) {
                  case 0:
                    _context8.prev = 0;
                    hoja = context.workbook.worksheets.getActiveWorksheet();
                    tabla = hoja.tables.getItem(nombreTabla);
                    encabezadosRange = tabla.getHeaderRowRange();
                    encabezadosRange.load("values");
                    _context8.next = 7;
                    return context.sync();
                  case 7:
                    encabezados = encabezadosRange.values[0];
                    select = document.getElementById(idSelect);
                    if (select) {
                      _context8.next = 11;
                      break;
                    }
                    return _context8.abrupt("return");
                  case 11:
                    if (select) {
                      select.disabled = false;
                    }
                    boton = document.getElementById("guardar");
                    if (boton) boton.disabled = false;
                    // Limpiar opciones actuales
                    select.innerHTML = "";

                    // Opción por defecto
                    opcionDefecto = document.createElement("option");
                    opcionDefecto.value = "";
                    opcionDefecto.textContent = "-- Selecciona una vista --";
                    select.appendChild(opcionDefecto);

                    // Añadir una opción por cada encabezado
                    encabezados.forEach(function (encabezado) {
                      var opcion = document.createElement("option");
                      opcion.value = encabezado;
                      opcion.textContent = encabezado;
                      select.appendChild(opcion);
                    });
                    _context8.next = 29;
                    break;
                  case 22:
                    _context8.prev = 22;
                    _context8.t0 = _context8["catch"](0);
                    _boton = document.getElementById("guardar");
                    if (_boton) _boton.disabled = true;
                    _select = document.getElementById(idSelect);
                    if (_select) _select.disabled = true;

                    // Aquí puedes llamar a la función que quieras
                    mostrarToast("No se encontr\xF3 la tabla ".concat(nombreTabla, ", por lo que no hay filtros predefinidos."), "#ffa94d");
                  case 29:
                  case "end":
                    return _context8.stop();
                }
              }, _callee7, null, [[0, 22]]);
            }));
            return function (_x16) {
              return _ref5.apply(this, arguments);
            };
          }());
        case 2:
        case "end":
          return _context9.stop();
      }
    }, _callee8);
  }));
  return _cargarOpcionesDesdeEncabezados.apply(this, arguments);
}
function aplicarVistaDesdeTFiltros(_x6, _x7) {
  return _aplicarVistaDesdeTFiltros.apply(this, arguments);
}
function _aplicarVistaDesdeTFiltros() {
  _aplicarVistaDesdeTFiltros = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee10(nombreCeldaTabla, idCombo) {
    var combo, vistaSeleccionada;
    return _regeneratorRuntime().wrap(function _callee10$(_context12) {
      while (1) switch (_context12.prev = _context12.next) {
        case 0:
          combo = document.getElementById(idCombo);
          vistaSeleccionada = combo.value;
          if (vistaSeleccionada) {
            _context12.next = 4;
            break;
          }
          return _context12.abrupt("return");
        case 4:
          _context12.next = 6;
          return Excel.run(/*#__PURE__*/function () {
            var _ref6 = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee9(context) {
              var workbook, namedCell, namedRange, nombreTablaPrincipal, hojaActiva, tablaPrincipal, columnas, hojas, tablaFiltros, _iterator2, _step2, hoja, posibleTabla, encabezadosRange, encabezadosFiltros, indiceVista, columnaVista, encabezadosAMostrar, botones, _iterator3, _step3, _loop2;
              return _regeneratorRuntime().wrap(function _callee9$(_context11) {
                while (1) switch (_context11.prev = _context11.next) {
                  case 0:
                    workbook = context.workbook; // 1. Obtener nombre de la tabla principal desde la celda con nombre definido
                    namedCell = workbook.names.getItem(nombreCeldaTabla);
                    namedRange = namedCell.getRange();
                    namedRange.load("values");
                    _context11.next = 6;
                    return context.sync();
                  case 6:
                    nombreTablaPrincipal = namedRange.values[0][0]; // 2. Obtener la tabla principal desde la hoja activa
                    hojaActiva = workbook.worksheets.getActiveWorksheet();
                    tablaPrincipal = hojaActiva.tables.getItem(nombreTablaPrincipal);
                    columnas = tablaPrincipal.columns;
                    columnas.load("items/name");
                    _context11.next = 13;
                    return context.sync();
                  case 13:
                    // 3. Buscar la tabla T_Filtros en cualquier hoja
                    hojas = workbook.worksheets;
                    hojas.load("items/name");
                    _context11.next = 17;
                    return context.sync();
                  case 17:
                    tablaFiltros = null;
                    _iterator2 = _createForOfIteratorHelper(hojas.items);
                    _context11.prev = 19;
                    _iterator2.s();
                  case 21:
                    if ((_step2 = _iterator2.n()).done) {
                      _context11.next = 37;
                      break;
                    }
                    hoja = _step2.value;
                    _context11.prev = 23;
                    posibleTabla = hoja.tables.getItem("T_Filtros");
                    posibleTabla.load("name");
                    _context11.next = 28;
                    return context.sync();
                  case 28:
                    tablaFiltros = posibleTabla;
                    return _context11.abrupt("break", 37);
                  case 32:
                    _context11.prev = 32;
                    _context11.t0 = _context11["catch"](23);
                    return _context11.abrupt("continue", 35);
                  case 35:
                    _context11.next = 21;
                    break;
                  case 37:
                    _context11.next = 42;
                    break;
                  case 39:
                    _context11.prev = 39;
                    _context11.t1 = _context11["catch"](19);
                    _iterator2.e(_context11.t1);
                  case 42:
                    _context11.prev = 42;
                    _iterator2.f();
                    return _context11.finish(42);
                  case 45:
                    if (tablaFiltros) {
                      _context11.next = 48;
                      break;
                    }
                    console.error("No se encontró la tabla T_Filtros en el libro.");
                    return _context11.abrupt("return");
                  case 48:
                    // 4. Obtener encabezados de T_Filtros
                    encabezadosRange = tablaFiltros.getHeaderRowRange();
                    encabezadosRange.load("values");
                    _context11.next = 52;
                    return context.sync();
                  case 52:
                    encabezadosFiltros = encabezadosRange.values[0].map(function (h) {
                      return h.toString().trim();
                    });
                    indiceVista = encabezadosFiltros.indexOf(vistaSeleccionada);
                    if (!(indiceVista === -1)) {
                      _context11.next = 57;
                      break;
                    }
                    console.warn("La vista seleccionada no existe en T_Filtros.");
                    return _context11.abrupt("return");
                  case 57:
                    // 5. Obtener los valores de la columna de la vista seleccionada
                    columnaVista = tablaFiltros.columns.getItemAt(indiceVista).getDataBodyRange();
                    columnaVista.load("values");
                    _context11.next = 61;
                    return context.sync();
                  case 61:
                    encabezadosAMostrar = columnaVista.values.map(function (fila) {
                      var _fila$;
                      return (_fila$ = fila[0]) === null || _fila$ === void 0 ? void 0 : _fila$.toString().trim().toLowerCase();
                    }).filter(Boolean); // 6. Mostrar/ocultar columnas y pintar botones
                    botones = document.querySelectorAll("#buttonsContainer button");
                    _iterator3 = _createForOfIteratorHelper(columnas.items);
                    _context11.prev = 64;
                    _loop2 = /*#__PURE__*/_regeneratorRuntime().mark(function _loop2() {
                      var col, nombreCol, nombreColLower, mostrar, rango;
                      return _regeneratorRuntime().wrap(function _loop2$(_context10) {
                        while (1) switch (_context10.prev = _context10.next) {
                          case 0:
                            col = _step3.value;
                            nombreCol = col.name.trim();
                            nombreColLower = nombreCol.toLowerCase();
                            mostrar = encabezadosAMostrar.includes(nombreColLower); // Mostrar u ocultar columna
                            rango = col.getRange();
                            rango.columnHidden = !mostrar;

                            // Actualizar color del botón
                            botones.forEach(function (btn) {
                              var _btn$textContent2;
                              if (((_btn$textContent2 = btn.textContent) === null || _btn$textContent2 === void 0 ? void 0 : _btn$textContent2.trim().toLowerCase()) === nombreColLower) {
                                btn.style.backgroundColor = mostrar ? "#cce5cc" : "#f5b5b5";
                              }
                            });
                          case 7:
                          case "end":
                            return _context10.stop();
                        }
                      }, _loop2);
                    });
                    _iterator3.s();
                  case 67:
                    if ((_step3 = _iterator3.n()).done) {
                      _context11.next = 71;
                      break;
                    }
                    return _context11.delegateYield(_loop2(), "t2", 69);
                  case 69:
                    _context11.next = 67;
                    break;
                  case 71:
                    _context11.next = 76;
                    break;
                  case 73:
                    _context11.prev = 73;
                    _context11.t3 = _context11["catch"](64);
                    _iterator3.e(_context11.t3);
                  case 76:
                    _context11.prev = 76;
                    _iterator3.f();
                    return _context11.finish(76);
                  case 79:
                    _context11.next = 81;
                    return context.sync();
                  case 81:
                  case "end":
                    return _context11.stop();
                }
              }, _callee9, null, [[19, 39, 42, 45], [23, 32], [64, 73, 76, 79]]);
            }));
            return function (_x17) {
              return _ref6.apply(this, arguments);
            };
          }());
        case 6:
        case "end":
          return _context12.stop();
      }
    }, _callee10);
  }));
  return _aplicarVistaDesdeTFiltros.apply(this, arguments);
}
function guardarVistaSeleccionada(_x8, _x9, _x10) {
  return _guardarVistaSeleccionada.apply(this, arguments);
}
function _guardarVistaSeleccionada() {
  _guardarVistaSeleccionada = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee12(idCombo, idBotones, nombreTablaFiltros) {
    var vista, encabezadosVisibles;
    return _regeneratorRuntime().wrap(function _callee12$(_context14) {
      while (1) switch (_context14.prev = _context14.next) {
        case 0:
          vista = getVistaSeleccionadaDesdeCombo(idCombo);
          if (vista) {
            _context14.next = 4;
            break;
          }
          mostrarToast("❌ Selecciona alguna vista para poder guardar", "#f28b94"); // rojo claro
          return _context14.abrupt("return");
        case 4:
          encabezadosVisibles = getEncabezadosVisiblesDesdeBotones(idBotones);
          if (!(encabezadosVisibles.length === 0)) {
            _context14.next = 8;
            break;
          }
          alert("No hay botones en verde. Nada que guardar.");
          return _context14.abrupt("return");
        case 8:
          _context14.next = 10;
          return Excel.run(/*#__PURE__*/function () {
            var _ref7 = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee11(context) {
              var workbook, hojas, tabla, hojaTabla, _iterator4, _step4, hoja, t, encabezadosRange, encabezados, indiceVista, columna, numFilas, celdasVacias, filasActualesResult, filasActuales, filasNecesarias, filasFaltantes, i;
              return _regeneratorRuntime().wrap(function _callee11$(_context13) {
                while (1) switch (_context13.prev = _context13.next) {
                  case 0:
                    workbook = context.workbook;
                    hojas = workbook.worksheets;
                    hojas.load("items/name");
                    _context13.next = 5;
                    return context.sync();
                  case 5:
                    tabla = null;
                    hojaTabla = null;
                    _iterator4 = _createForOfIteratorHelper(hojas.items);
                    _context13.prev = 8;
                    _iterator4.s();
                  case 10:
                    if ((_step4 = _iterator4.n()).done) {
                      _context13.next = 27;
                      break;
                    }
                    hoja = _step4.value;
                    _context13.prev = 12;
                    t = hoja.tables.getItem(nombreTablaFiltros);
                    t.load("name");
                    _context13.next = 17;
                    return context.sync();
                  case 17:
                    tabla = t;
                    hojaTabla = hoja;
                    return _context13.abrupt("break", 27);
                  case 22:
                    _context13.prev = 22;
                    _context13.t0 = _context13["catch"](12);
                    return _context13.abrupt("continue", 25);
                  case 25:
                    _context13.next = 10;
                    break;
                  case 27:
                    _context13.next = 32;
                    break;
                  case 29:
                    _context13.prev = 29;
                    _context13.t1 = _context13["catch"](8);
                    _iterator4.e(_context13.t1);
                  case 32:
                    _context13.prev = 32;
                    _iterator4.f();
                    return _context13.finish(32);
                  case 35:
                    if (!(!tabla || !hojaTabla)) {
                      _context13.next = 38;
                      break;
                    }
                    console.error("No se encontr\xF3 la tabla ".concat(nombreTablaFiltros, "."));
                    return _context13.abrupt("return");
                  case 38:
                    // Buscar índice de columna correspondiente a la vista
                    encabezadosRange = tabla.getHeaderRowRange();
                    encabezadosRange.load("values");
                    _context13.next = 42;
                    return context.sync();
                  case 42:
                    encabezados = encabezadosRange.values[0].map(function (h) {
                      return h.toString().trim();
                    });
                    indiceVista = encabezados.indexOf(vista);
                    if (!(indiceVista === -1)) {
                      _context13.next = 47;
                      break;
                    }
                    console.warn("La vista seleccionada no existe en T_Filtros.");
                    return _context13.abrupt("return");
                  case 47:
                    // Obtener rango actual de la columna para limpiar
                    columna = tabla.columns.getItemAt(indiceVista).getDataBodyRange();
                    columna.load("rowCount");
                    _context13.next = 51;
                    return context.sync();
                  case 51:
                    numFilas = columna.rowCount; // Limpiar contenido actual
                    if (!(numFilas > 0)) {
                      _context13.next = 57;
                      break;
                    }
                    celdasVacias = Array(numFilas).fill([""]);
                    columna.values = celdasVacias;
                    _context13.next = 57;
                    return context.sync();
                  case 57:
                    // Verificar si hay suficientes filas en la tabla para los nuevos valores
                    filasActualesResult = columna.rowCount;
                    _context13.next = 60;
                    return context.sync();
                  case 60:
                    filasActuales = filasActualesResult;
                    filasNecesarias = encabezadosVisibles.length;
                    if (!(filasActuales < filasNecesarias)) {
                      _context13.next = 67;
                      break;
                    }
                    filasFaltantes = filasNecesarias - filasActuales;
                    for (i = 0; i < filasFaltantes; i++) {
                      tabla.rows.add(null, [[]]); // Agrega filas vacías
                    }
                    _context13.next = 67;
                    return context.sync();
                  case 67:
                    // ⚠️ IMPORTANTE: volver a cargar el rango con las filas nuevas
                    columna = tabla.columns.getItemAt(indiceVista).getDataBodyRange();
                    _context13.next = 70;
                    return context.sync();
                  case 70:
                    // Asignar los nuevos valores fila por fila (evita error de dimensión)
                    encabezadosVisibles.forEach(function (valor, i) {
                      columna.getCell(i, 0).values = [[valor]];
                    });
                    _context13.next = 73;
                    return context.sync();
                  case 73:
                    mostrarToast("\u2705 Guardado correctamente en la vista \"".concat(vista, "\""));
                  case 74:
                  case "end":
                    return _context13.stop();
                }
              }, _callee11, null, [[8, 29, 32, 35], [12, 22]]);
            }));
            return function (_x18) {
              return _ref7.apply(this, arguments);
            };
          }());
        case 10:
        case "end":
          return _context14.stop();
      }
    }, _callee12);
  }));
  return _guardarVistaSeleccionada.apply(this, arguments);
}
function getVistaSeleccionadaDesdeCombo(idCombo) {
  var _combo$value;
  var combo = document.getElementById(idCombo);
  return (combo === null || combo === void 0 || (_combo$value = combo.value) === null || _combo$value === void 0 ? void 0 : _combo$value.trim()) || null;
}
function getEncabezadosVisiblesDesdeBotones(containerId) {
  var botones = document.querySelectorAll("#".concat(containerId, " button"));
  return Array.from(botones).filter(function (btn) {
    var color = btn.style.backgroundColor.replace(/\s/g, "").toLowerCase();
    return color === "rgb(204,229,204)" || color === "#cce5cc";
  }).map(function (btn) {
    var _btn$textContent;
    return ((_btn$textContent = btn.textContent) === null || _btn$textContent === void 0 ? void 0 : _btn$textContent.trim()) || "";
  }).filter(Boolean);
}
function mostrarToast(mensaje) {
  var colorFondo = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "#b6e7b0";
  var toast = document.getElementById("toastGuardado");
  if (!toast) return;
  toast.textContent = mensaje;
  toast.style.backgroundColor = colorFondo;
  toast.style.display = "block";
  setTimeout(function () {
    toast.style.display = "none";
  }, 5000); // Puedes ajustar la duración si quieres
}
function reordenarColumnasEnExcel(_x11, _x12) {
  return _reordenarColumnasEnExcel.apply(this, arguments);
}
function _reordenarColumnasEnExcel() {
  _reordenarColumnasEnExcel = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee14(nombreTabla, nuevoOrden) {
    var mensaje;
    return _regeneratorRuntime().wrap(function _callee14$(_context16) {
      while (1) switch (_context16.prev = _context16.next) {
        case 0:
          _context16.prev = 0;
          _context16.next = 3;
          return Excel.run(/*#__PURE__*/function () {
            var _ref8 = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee13(context) {
              var hoja, tabla, columnas, ordenActual, i, nombreDeseado, posicionActual, columna, rango, datos, filasOrigen, columnasOrigen, nuevaColumna, destino, filasDestino, filasFaltantes, j, destinoFinal, filasFinal, columnasFinal, datosAjustados, f, _datos$f$, _datos$f, valor, indexAEliminar;
              return _regeneratorRuntime().wrap(function _callee13$(_context15) {
                while (1) switch (_context15.prev = _context15.next) {
                  case 0:
                    hoja = context.workbook.worksheets.getActiveWorksheet();
                    tabla = hoja.tables.getItem(nombreTabla);
                    columnas = tabla.columns;
                    columnas.load("items/name");
                    _context15.next = 6;
                    return context.sync();
                  case 6:
                    ordenActual = columnas.items.map(function (col) {
                      return col.name;
                    });
                    i = 0;
                  case 8:
                    if (!(i < nuevoOrden.length)) {
                      _context15.next = 61;
                      break;
                    }
                    nombreDeseado = nuevoOrden[i];
                    posicionActual = ordenActual.indexOf(nombreDeseado);
                    if (!(posicionActual === -1 || posicionActual === i)) {
                      _context15.next = 13;
                      break;
                    }
                    return _context15.abrupt("continue", 58);
                  case 13:
                    columna = tabla.columns.getItemAt(posicionActual);
                    rango = columna.getDataBodyRange();
                    rango.load(["values", "rowCount", "columnCount"]);
                    _context15.next = 18;
                    return context.sync();
                  case 18:
                    datos = rango.values || [];
                    filasOrigen = rango.rowCount;
                    columnasOrigen = rango.columnCount;
                    if (filasOrigen === 0 || datos.length === 0) {
                      datos = [[""]];
                      filasOrigen = 1;
                      mostrarToast("\u2139\uFE0F Columna \"".concat(nombreDeseado, "\" estaba vac\xEDa. Se a\xF1adi\xF3 fila vac\xEDa."), "#ffeeba");
                    }

                    // Insertar nueva columna vacía en la posición deseada
                    tabla.columns.add(i, [[nombreDeseado]]);
                    _context15.next = 25;
                    return context.sync();
                  case 25:
                    nuevaColumna = tabla.columns.getItemAt(i);
                    destino = nuevaColumna.getDataBodyRange();
                    destino.load(["rowCount", "columnCount"]);
                    _context15.next = 30;
                    return context.sync();
                  case 30:
                    filasDestino = destino.rowCount;
                    if (!(filasDestino < filasOrigen)) {
                      _context15.next = 36;
                      break;
                    }
                    filasFaltantes = filasOrigen - filasDestino;
                    for (j = 0; j < filasFaltantes; j++) {
                      tabla.rows.add(null, [[]]);
                    }
                    _context15.next = 36;
                    return context.sync();
                  case 36:
                    destinoFinal = nuevaColumna.getDataBodyRange();
                    destinoFinal.load(["rowCount", "columnCount"]);
                    _context15.next = 40;
                    return context.sync();
                  case 40:
                    filasFinal = destinoFinal.rowCount;
                    columnasFinal = destinoFinal.columnCount;
                    if (!(columnasFinal === 1)) {
                      _context15.next = 50;
                      break;
                    }
                    datosAjustados = [];
                    for (f = 0; f < filasFinal; f++) {
                      valor = (_datos$f$ = (_datos$f = datos[f]) === null || _datos$f === void 0 ? void 0 : _datos$f[0]) !== null && _datos$f$ !== void 0 ? _datos$f$ : "";
                      if (Array.isArray(valor)) valor = valor[0];
                      datosAjustados.push([valor]);
                    }
                    destinoFinal.values = datosAjustados;
                    _context15.next = 48;
                    return context.sync();
                  case 48:
                    _context15.next = 52;
                    break;
                  case 50:
                    mostrarToast("\u274C No se pudo copiar \"".concat(nombreDeseado, "\": columnas incompatibles."), "#f28b94");
                    return _context15.abrupt("continue", 58);
                  case 52:
                    indexAEliminar = posicionActual >= i ? posicionActual + 1 : posicionActual;
                    tabla.columns.getItemAt(indexAEliminar).delete();
                    _context15.next = 56;
                    return context.sync();
                  case 56:
                    ordenActual.splice(posicionActual, 1);
                    ordenActual.splice(i, 0, nombreDeseado);
                  case 58:
                    i++;
                    _context15.next = 8;
                    break;
                  case 61:
                    mostrarToast("✅ Columnas reordenadas correctamente");
                  case 62:
                  case "end":
                    return _context15.stop();
                }
              }, _callee13);
            }));
            return function (_x19) {
              return _ref8.apply(this, arguments);
            };
          }());
        case 3:
          _context16.next = 10;
          break;
        case 5:
          _context16.prev = 5;
          _context16.t0 = _context16["catch"](0);
          mensaje = (_context16.t0 === null || _context16.t0 === void 0 ? void 0 : _context16.t0.message) || "❌ Error desconocido al reordenar columnas.";
          mostrarToast("\u274C ERROR: ".concat(mensaje), "#f28b94");
          console.error(_context16.t0);
        case 10:
        case "end":
          return _context16.stop();
      }
    }, _callee14, null, [[0, 5]]);
  }));
  return _reordenarColumnasEnExcel.apply(this, arguments);
}
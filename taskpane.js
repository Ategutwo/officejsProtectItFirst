/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "8d768f65702f2137206f.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript && document.currentScript.tagName.toUpperCase() === 'SCRIPT')
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/^blob:/, "").replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = document.baseURI || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   run: function() { return /* binding */ run; }
/* harmony export */ });
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _regeneratorRuntime() { "use strict"; /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/facebook/regenerator/blob/main/LICENSE */ _regeneratorRuntime = function _regeneratorRuntime() { return e; }; var t, e = {}, r = Object.prototype, n = r.hasOwnProperty, o = Object.defineProperty || function (t, e, r) { t[e] = r.value; }, i = "function" == typeof Symbol ? Symbol : {}, a = i.iterator || "@@iterator", c = i.asyncIterator || "@@asyncIterator", u = i.toStringTag || "@@toStringTag"; function define(t, e, r) { return Object.defineProperty(t, e, { value: r, enumerable: !0, configurable: !0, writable: !0 }), t[e]; } try { define({}, ""); } catch (t) { define = function define(t, e, r) { return t[e] = r; }; } function wrap(t, e, r, n) { var i = e && e.prototype instanceof Generator ? e : Generator, a = Object.create(i.prototype), c = new Context(n || []); return o(a, "_invoke", { value: makeInvokeMethod(t, r, c) }), a; } function tryCatch(t, e, r) { try { return { type: "normal", arg: t.call(e, r) }; } catch (t) { return { type: "throw", arg: t }; } } e.wrap = wrap; var h = "suspendedStart", l = "suspendedYield", f = "executing", s = "completed", y = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} var p = {}; define(p, a, function () { return this; }); var d = Object.getPrototypeOf, v = d && d(d(values([]))); v && v !== r && n.call(v, a) && (p = v); var g = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(p); function defineIteratorMethods(t) { ["next", "throw", "return"].forEach(function (e) { define(t, e, function (t) { return this._invoke(e, t); }); }); } function AsyncIterator(t, e) { function invoke(r, o, i, a) { var c = tryCatch(t[r], t, o); if ("throw" !== c.type) { var u = c.arg, h = u.value; return h && "object" == _typeof(h) && n.call(h, "__await") ? e.resolve(h.__await).then(function (t) { invoke("next", t, i, a); }, function (t) { invoke("throw", t, i, a); }) : e.resolve(h).then(function (t) { u.value = t, i(u); }, function (t) { return invoke("throw", t, i, a); }); } a(c.arg); } var r; o(this, "_invoke", { value: function value(t, n) { function callInvokeWithMethodAndArg() { return new e(function (e, r) { invoke(t, n, e, r); }); } return r = r ? r.then(callInvokeWithMethodAndArg, callInvokeWithMethodAndArg) : callInvokeWithMethodAndArg(); } }); } function makeInvokeMethod(e, r, n) { var o = h; return function (i, a) { if (o === f) throw Error("Generator is already running"); if (o === s) { if ("throw" === i) throw a; return { value: t, done: !0 }; } for (n.method = i, n.arg = a;;) { var c = n.delegate; if (c) { var u = maybeInvokeDelegate(c, n); if (u) { if (u === y) continue; return u; } } if ("next" === n.method) n.sent = n._sent = n.arg;else if ("throw" === n.method) { if (o === h) throw o = s, n.arg; n.dispatchException(n.arg); } else "return" === n.method && n.abrupt("return", n.arg); o = f; var p = tryCatch(e, r, n); if ("normal" === p.type) { if (o = n.done ? s : l, p.arg === y) continue; return { value: p.arg, done: n.done }; } "throw" === p.type && (o = s, n.method = "throw", n.arg = p.arg); } }; } function maybeInvokeDelegate(e, r) { var n = r.method, o = e.iterator[n]; if (o === t) return r.delegate = null, "throw" === n && e.iterator.return && (r.method = "return", r.arg = t, maybeInvokeDelegate(e, r), "throw" === r.method) || "return" !== n && (r.method = "throw", r.arg = new TypeError("The iterator does not provide a '" + n + "' method")), y; var i = tryCatch(o, e.iterator, r.arg); if ("throw" === i.type) return r.method = "throw", r.arg = i.arg, r.delegate = null, y; var a = i.arg; return a ? a.done ? (r[e.resultName] = a.value, r.next = e.nextLoc, "return" !== r.method && (r.method = "next", r.arg = t), r.delegate = null, y) : a : (r.method = "throw", r.arg = new TypeError("iterator result is not an object"), r.delegate = null, y); } function pushTryEntry(t) { var e = { tryLoc: t[0] }; 1 in t && (e.catchLoc = t[1]), 2 in t && (e.finallyLoc = t[2], e.afterLoc = t[3]), this.tryEntries.push(e); } function resetTryEntry(t) { var e = t.completion || {}; e.type = "normal", delete e.arg, t.completion = e; } function Context(t) { this.tryEntries = [{ tryLoc: "root" }], t.forEach(pushTryEntry, this), this.reset(!0); } function values(e) { if (e || "" === e) { var r = e[a]; if (r) return r.call(e); if ("function" == typeof e.next) return e; if (!isNaN(e.length)) { var o = -1, i = function next() { for (; ++o < e.length;) if (n.call(e, o)) return next.value = e[o], next.done = !1, next; return next.value = t, next.done = !0, next; }; return i.next = i; } } throw new TypeError(_typeof(e) + " is not iterable"); } return GeneratorFunction.prototype = GeneratorFunctionPrototype, o(g, "constructor", { value: GeneratorFunctionPrototype, configurable: !0 }), o(GeneratorFunctionPrototype, "constructor", { value: GeneratorFunction, configurable: !0 }), GeneratorFunction.displayName = define(GeneratorFunctionPrototype, u, "GeneratorFunction"), e.isGeneratorFunction = function (t) { var e = "function" == typeof t && t.constructor; return !!e && (e === GeneratorFunction || "GeneratorFunction" === (e.displayName || e.name)); }, e.mark = function (t) { return Object.setPrototypeOf ? Object.setPrototypeOf(t, GeneratorFunctionPrototype) : (t.__proto__ = GeneratorFunctionPrototype, define(t, u, "GeneratorFunction")), t.prototype = Object.create(g), t; }, e.awrap = function (t) { return { __await: t }; }, defineIteratorMethods(AsyncIterator.prototype), define(AsyncIterator.prototype, c, function () { return this; }), e.AsyncIterator = AsyncIterator, e.async = function (t, r, n, o, i) { void 0 === i && (i = Promise); var a = new AsyncIterator(wrap(t, r, n, o), i); return e.isGeneratorFunction(r) ? a : a.next().then(function (t) { return t.done ? t.value : a.next(); }); }, defineIteratorMethods(g), define(g, u, "Generator"), define(g, a, function () { return this; }), define(g, "toString", function () { return "[object Generator]"; }), e.keys = function (t) { var e = Object(t), r = []; for (var n in e) r.push(n); return r.reverse(), function next() { for (; r.length;) { var t = r.pop(); if (t in e) return next.value = t, next.done = !1, next; } return next.done = !0, next; }; }, e.values = values, Context.prototype = { constructor: Context, reset: function reset(e) { if (this.prev = 0, this.next = 0, this.sent = this._sent = t, this.done = !1, this.delegate = null, this.method = "next", this.arg = t, this.tryEntries.forEach(resetTryEntry), !e) for (var r in this) "t" === r.charAt(0) && n.call(this, r) && !isNaN(+r.slice(1)) && (this[r] = t); }, stop: function stop() { this.done = !0; var t = this.tryEntries[0].completion; if ("throw" === t.type) throw t.arg; return this.rval; }, dispatchException: function dispatchException(e) { if (this.done) throw e; var r = this; function handle(n, o) { return a.type = "throw", a.arg = e, r.next = n, o && (r.method = "next", r.arg = t), !!o; } for (var o = this.tryEntries.length - 1; o >= 0; --o) { var i = this.tryEntries[o], a = i.completion; if ("root" === i.tryLoc) return handle("end"); if (i.tryLoc <= this.prev) { var c = n.call(i, "catchLoc"), u = n.call(i, "finallyLoc"); if (c && u) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } else if (c) { if (this.prev < i.catchLoc) return handle(i.catchLoc, !0); } else { if (!u) throw Error("try statement without catch or finally"); if (this.prev < i.finallyLoc) return handle(i.finallyLoc); } } } }, abrupt: function abrupt(t, e) { for (var r = this.tryEntries.length - 1; r >= 0; --r) { var o = this.tryEntries[r]; if (o.tryLoc <= this.prev && n.call(o, "finallyLoc") && this.prev < o.finallyLoc) { var i = o; break; } } i && ("break" === t || "continue" === t) && i.tryLoc <= e && e <= i.finallyLoc && (i = null); var a = i ? i.completion : {}; return a.type = t, a.arg = e, i ? (this.method = "next", this.next = i.finallyLoc, y) : this.complete(a); }, complete: function complete(t, e) { if ("throw" === t.type) throw t.arg; return "break" === t.type || "continue" === t.type ? this.next = t.arg : "return" === t.type ? (this.rval = this.arg = t.arg, this.method = "return", this.next = "end") : "normal" === t.type && e && (this.next = e), y; }, finish: function finish(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.finallyLoc === t) return this.complete(r.completion, r.afterLoc), resetTryEntry(r), y; } }, catch: function _catch(t) { for (var e = this.tryEntries.length - 1; e >= 0; --e) { var r = this.tryEntries[e]; if (r.tryLoc === t) { var n = r.completion; if ("throw" === n.type) { var o = n.arg; resetTryEntry(r); } return o; } } throw Error("illegal catch attempt"); }, delegateYield: function delegateYield(e, r, n) { return this.delegate = { iterator: values(e), resultName: r, nextLoc: n }, "next" === this.method && (this.arg = t), y; } }, e; }
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t.return || t.return(); } finally { if (u) throw o; } } }; }
function _slicedToArray(r, e) { return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest(); }
function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t.return && (u = t.return(), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function _arrayWithHoles(r) { if (Array.isArray(r)) return r; }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = run;
  }
});
function run() {
  return _run.apply(this, arguments);
}

// ─── 1. Define your package ────────────────────────────────────────────────
// const packages = {
//   emk1: {
//     packagePrice: 100,       // one‐time package sale price
//     drugs: [
//       { id: 'a', price: 10,  replenishDays: 300 },  // auto-replenishable every 300d
//       { id: 'b', price: 15,  replenishDays: 200 },  
//       { id: 'c', price: 5,   replenishDays: 550 },
//       { id: 'd', price: 8 }   // no replenishDays → non-replenishable
//     ]
//   }
// };

// // ─── 2. Your past + future sales maps ─────────────────────────────────────
// // format: { "YYYY-MM": unitsSold }
// const salesHistory = {
//   '2022-01': 10,
//   '2022-02': 12,
//   // … all of 2022, 2023, 2024 …
//   '2025-01': 8    // Jan 2025 sales
// };
// const projectedSales = generateProjections('2025-02', '2035-03', 5);

// ─── 3. Revenue calculator ────────────────────────────────────────────────
// --- Step 1: Monthly base revenue
function _run() {
  _run = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee2() {
    return _regeneratorRuntime().wrap(function _callee2$(_context2) {
      while (1) switch (_context2.prev = _context2.next) {
        case 0:
          _context2.prev = 0;
          _context2.next = 3;
          return Excel.run(/*#__PURE__*/function () {
            var _ref3 = _asyncToGenerator(/*#__PURE__*/_regeneratorRuntime().mark(function _callee(context) {
              var ws, packageDetails, packageDetailsRange, usedRange, drugsExpirationPredictions, wsAutoReplenishMedGroups, wsRevenuePredictions, lastRow, data, packageDetailsData, medsObj, emkDetails, wsNewKit, newKitsLastRow, newKitsLastRowIndex, dataRange, newKitData, salesHistory, calculatedKitData, newKitDrugPredictions, updatedDrugData, baseMap, forecastMap, drugDataMap, rangeAutoReplenishMedGroups, autoReplenish, allMonths, finalRevenueForecast, _iterator2, _step2, month, newkit, auto, drugData, totalRevenue;
              return _regeneratorRuntime().wrap(function _callee$(_context) {
                while (1) switch (_context.prev = _context.next) {
                  case 0:
                    ws = context.workbook.worksheets.getItem("DrugDetails");
                    packageDetails = context.workbook.worksheets.getItem("packageDistribution");
                    packageDetailsRange = packageDetails.getRange("A2:D7");
                    usedRange = ws.getUsedRange().getLastRow();
                    drugsExpirationPredictions = context.workbook.worksheets.getItem("Drug Replenish Dates(New Kits)");
                    wsAutoReplenishMedGroups = context.workbook.worksheets.getItem("auto_replenish_med_groups");
                    wsRevenuePredictions = context.workbook.worksheets.getItem("Revenue Prediction");
                    wsRevenuePredictions.getRangeByIndexes(1, 0, 10000, 50).clear(Excel.ClearApplyTo.contents);
                    drugsExpirationPredictions.getRangeByIndexes(1, 0, 10000, 50).clear(Excel.ClearApplyTo.contents);
                    //Get the Details  
                    usedRange.load("rowIndex");
                    _context.next = 12;
                    return context.sync();
                  case 12:
                    lastRow = usedRange.rowIndex;
                    data = ws.getRange("B".concat(1, ":O", lastRow + 1));
                    data.load("values");
                    packageDetailsRange.load("values");
                    _context.next = 18;
                    return context.sync();
                  case 18:
                    packageDetailsData = packageDetailsRange.values;
                    medsObj = {};
                    emkDetails = {}; //Get the drug details
                    console.log(data.values);
                    data.values.forEach(function (row) {
                      medsObj[row[0]] = {
                        totalUnitCost: row[3],
                        laCarte: row[4],
                        includedInPackages: [],
                        shelfLife: row[7]
                      };
                      for (var i = 8; i <= 13; i++) {
                        if (row[i].toString().trim() !== "") {
                          medsObj[row[0]].includedInPackages.push(data.values[0][i]);
                        }
                      }
                    });
                    packageDetailsData.forEach(function (row) {
                      //Create the emk objecst
                      emkDetails[row[0]] = {
                        retailPrice: row[1],
                        newKitShares: row[2],
                        purchasePrice: row[3],
                        drugs: []
                      };
                    });
                    // console.log(medsObj,emkDetails)

                    //Get the New Kit Data
                    wsNewKit = context.workbook.worksheets.getItem("New Kit Data");
                    newKitsLastRow = wsNewKit.getUsedRange().getLastRow();
                    newKitsLastRow.load("rowIndex");
                    _context.next = 29;
                    return context.sync();
                  case 29:
                    newKitsLastRowIndex = newKitsLastRow.rowIndex;
                    dataRange = wsNewKit.getRange("A2:B".concat(newKitsLastRowIndex + 1));
                    dataRange.load("values");
                    _context.next = 34;
                    return context.sync();
                  case 34:
                    newKitData = dataRange.values;
                    salesHistory = {}; //Get the Kit Revenue for each Kit and total Revenue
                    calculatedKitData = newKitData.map(function (row) {
                      salesHistory[formatDate(excelSerialDateToJSDate(row[0]))] = row[1];
                      var numberOfKits = row[1];
                      var EMK1 = Math.floor(emkDetails["EMK1"].newKitShares * numberOfKits) * emkDetails["EMK1"].retailPrice;
                      var EMK5 = Math.floor(emkDetails["EMK5"].newKitShares * numberOfKits) * emkDetails["EMK5"].retailPrice;
                      var EMK10 = Math.floor(emkDetails["EMK10"].newKitShares * numberOfKits) * emkDetails["EMK10"].retailPrice;
                      var EMK15 = Math.floor(emkDetails["EMK15"].newKitShares * numberOfKits) * emkDetails["EMK15"].retailPrice;
                      var EMK1Mini = Math.floor(emkDetails["EMK1-Mini"].newKitShares * numberOfKits) * emkDetails["EMK1-Mini"].retailPrice;
                      var EMK10Mini = Math.floor(emkDetails["EMK10-Mini"].newKitShares * numberOfKits) * emkDetails["EMK10-Mini"].retailPrice;
                      return [row[0], row[1], EMK1 + EMK5 + EMK10 + EMK15 + EMK1Mini + EMK10Mini, "", EMK1, EMK5, EMK10, EMK15, EMK1Mini, EMK10Mini];
                    }); //Add the Kit Revenue to the sheet
                    wsNewKit.getRange("A2:J" + (calculatedKitData.length + 1)).values = calculatedKitData;
                    //Add the total  Revenue to the sheet 
                    // const revenueLedger = calcRevenue(packages.emk1, salesHistory, projectedSales);
                    // console.log(revenueLedger);  
                    //Get the drugs that belong to each Kit 
                    data.values.forEach(function (row) {
                      row[8] === "X" ? emkDetails["EMK1"]["drugs"].push(row[0]) : "";
                      row[9] === "X" ? emkDetails["EMK5"]["drugs"].push(row[0]) : "";
                      row[10] === "X" ? emkDetails["EMK10"]["drugs"].push(row[0]) : "";
                      row[11] === "X" ? emkDetails["EMK15"]["drugs"].push(row[0]) : "";
                      row[12] === "X" ? emkDetails["EMK1-Mini"]["drugs"].push(row[0]) : "";
                      row[13] === "X" ? emkDetails["EMK10-Mini"]["drugs"].push(row[0]) : "";
                    });
                    //Creating calculation for all drugs per month
                    newKitDrugPredictions = [];
                    Object.keys(salesHistory).forEach(function (month) {
                      var totalKitAmount = salesHistory[month];
                      Object.keys(emkDetails).forEach(function (kit) {
                        var kitAmount = Math.floor(totalKitAmount * emkDetails[kit].newKitShares);
                        if (kitAmount < 1) return;
                        emkDetails[kit].drugs.forEach(function (drug) {
                          if (medsObj[drug].shelfLife == "" || medsObj[drug].shelfLife == "N/A") return;
                          newKitDrugPredictions.push([month, kit, drug, kitAmount, medsObj[drug].laCarte * kitAmount, medsObj[drug].shelfLife]);
                        });
                      });
                    });

                    //Adding Replenish Dates to the Drug Details
                    updatedDrugData = newKitDrugPredictions.map(function (row) {
                      var _row2 = _slicedToArray(row, 6),
                        date = _row2[0],
                        code = _row2[1],
                        description = _row2[2],
                        qty = _row2[3],
                        total = _row2[4],
                        expiryDays = _row2[5];
                      var _date$split$map = date.split("-").map(Number),
                        _date$split$map2 = _slicedToArray(_date$split$map, 2),
                        year = _date$split$map2[0],
                        month = _date$split$map2[1];
                      var baseDate = new Date(year, month - 1);
                      var replenishments = [];
                      for (var i = 1; i <= 10; i++) {
                        var expireDate = new Date(baseDate);
                        expireDate.setDate(expireDate.getDate() + expiryDays * i);
                        var expireYear = expireDate.getFullYear();
                        var expireMonth = String(expireDate.getMonth() + 1).padStart(2, '0');
                        replenishments.push("".concat(expireYear, "-").concat(expireMonth));
                      }
                      return [].concat(_toConsumableArray(row), replenishments);
                    });
                    drugsExpirationPredictions.getRangeByIndexes(1, 0, updatedDrugData.length, updatedDrugData[0].length).values = updatedDrugData;
                    // --- Step 5: Execute everything
                    baseMap = getBaseKitMap(calculatedKitData);
                    forecastMap = generateForecast("2025-05", 120, baseMap); // Plug in your generated updatedDrugData (with replenishment dates)
                    drugDataMap = applyDrugDataRevenue(forecastMap, updatedDrugData); //Get Autor replenish sheet data
                    rangeAutoReplenishMedGroups = wsAutoReplenishMedGroups.getRange("D2:F22011");
                    rangeAutoReplenishMedGroups.load("values");
                    _context.next = 50;
                    return context.sync();
                  case 50:
                    // Auto-replenish items (only applied once)
                    console.log(rangeAutoReplenishMedGroups.values.splice(0, 10));
                    autoReplenish = applyAutoReplenishOnce(forecastMap, rangeAutoReplenishMedGroups.values);
                    console.log(drugDataMap, baseMap, autoReplenish);
                    // 1. Combine all unique months

                    // --- Step 6: Final Output
                    // const finalRevenueForecast = Array.from(forecastMap.entries()).map(([month, revenue]) => [month, revenue]);
                    allMonths = new Set([].concat(_toConsumableArray(drugDataMap.keys()), _toConsumableArray(autoReplenish.keys()), _toConsumableArray(baseMap.keys()), _toConsumableArray(forecastMap.keys()))); // 2. Generate final forecast array
                    finalRevenueForecast = [];
                    _iterator2 = _createForOfIteratorHelper(_toConsumableArray(allMonths).sort());
                    try {
                      for (_iterator2.s(); !(_step2 = _iterator2.n()).done;) {
                        month = _step2.value;
                        newkit = baseMap.get(month) || 0;
                        auto = autoReplenish.get(month) || 0;
                        drugData = drugDataMap.get(month) || 0;
                        totalRevenue = newkit + auto + drugData;
                        finalRevenueForecast.push([month, totalRevenue, newkit, auto, drugData]);
                      }
                    } catch (err) {
                      _iterator2.e(err);
                    } finally {
                      _iterator2.f();
                    }
                    wsRevenuePredictions.getRangeByIndexes(1, 0, finalRevenueForecast.length, finalRevenueForecast[0].length).values = finalRevenueForecast;
                    console.table(finalRevenueForecast);
                    return _context.abrupt("return", context.sync());
                  case 60:
                  case "end":
                    return _context.stop();
                }
              }, _callee);
            }));
            return function (_x) {
              return _ref3.apply(this, arguments);
            };
          }());
        case 3:
          _context2.next = 8;
          break;
        case 5:
          _context2.prev = 5;
          _context2.t0 = _context2["catch"](0);
          console.error(_context2.t0);
        case 8:
        case "end":
          return _context2.stop();
      }
    }, _callee2, null, [[0, 5]]);
  }));
  return _run.apply(this, arguments);
}
function getBaseKitMap(baseKitRevenue) {
  var map = new Map();
  baseKitRevenue.forEach(function (_ref) {
    var _ref2 = _slicedToArray(_ref, 3),
      dateStr = _ref2[0],
      kitQuantity = _ref2[1],
      revenue = _ref2[2];
    var date = new excelSerialDateToJSDate(dateStr);
    var key = "".concat(date.getFullYear(), "-").concat(String(date.getMonth() + 1).padStart(2, '0'));
    map.set(key, revenue);
  });
  return map;
}

// --- Step 2: Forecast structure (June 2023 → May 2033)
function generateForecast() {
  var start = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : "2023-06";
  var months = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 120;
  var baseMap = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : new Map();
  var forecast = new Map();
  var _start$split$map = start.split("-").map(Number),
    _start$split$map2 = _slicedToArray(_start$split$map, 2),
    startYear = _start$split$map2[0],
    startMonth = _start$split$map2[1];
  var date = new Date(startYear, startMonth - 1);
  for (var i = 0; i < months; i++) {
    var year = date.getFullYear();
    var month = String(date.getMonth() + 1).padStart(2, '0');
    var key = "".concat(year, "-").concat(month);
    var baseRevenue = baseMap.get(key) || 0;
    forecast.set(key, baseRevenue);
    date.setMonth(date.getMonth() + 1);
  }
  return forecast;
}

// --- Step 3: Add drugData replenishment costs
function applyDrugDataRevenue(forecastMap, drugData) {
  var drugDataMap = new Map();
  var _iterator = _createForOfIteratorHelper(drugData),
    _step;
  try {
    var _loop = function _loop() {
      var row = _step.value;
      var total = parseFloat(row[4]);
      var replenishmentDates = row.slice(6);
      // dynamically added dates
      replenishmentDates.forEach(function (date) {
        if (forecastMap.has(date)) {
          forecastMap.set(date, forecastMap.get(date) + total);
          drugDataMap.set(date, drugDataMap.get(date) != undefined ? drugDataMap.get(date) + total : total);
        }
      });
    };
    for (_iterator.s(); !(_step = _iterator.n()).done;) {
      _loop();
    }
  } catch (err) {
    _iterator.e(err);
  } finally {
    _iterator.f();
  }
  return drugDataMap;
}

// --- Step 4: Add Auto Replenish (just once, at expiration date)
function applyAutoReplenishOnce(forecastMap, autoData) {
  var autoReplenish = new Map();
  autoData.forEach(function (row) {
    var _row = _slicedToArray(row, 3),
      expDate = _row[0],
      priceStr = _row[1],
      status = _row[2];
    if (status !== "Enabled") return;
    var price = typeof priceStr == "string" ? parseFloat(priceStr.replace("$", "")) : priceStr;

    // const [expMonth, , expYear] = expDate.split("/").map(Number);
    // const key = `${expYear}-${String(expMonth).padStart(2, '0')}`;
    var date = new excelSerialDateToJSDate(expDate);
    var key = "".concat(date.getFullYear(), "-").concat(String(date.getMonth() + 1).padStart(2, '0'));
    if (forecastMap.has(key)) {
      if (!isNaN(price)) forecastMap.set(key, forecastMap.get(key) + price);
      autoReplenish.set(key, autoReplenish.get(key) !== undefined ? autoReplenish.get(key) + price : price);
    }
  });
  return autoReplenish;
}

// // --- Step 5: Execute everything
// const baseMap = getBaseKitMap(baseKitRevenue);
// const forecastMap = generateForecast("2023-06", 120, baseMap);

// // Plug in your generated updatedDrugData (with replenishment dates)
// applyDrugDataRevenue(forecastMap, updatedDrugData);

// // Auto-replenish items (only applied once)
// applyAutoReplenishOnce(forecastMap, [
//   ["42", "Dental Depot", "Insta-Glucose", "2/28/2026", "$10.85", "Enabled"],
//   ["42", "Dental Depot", "Nitroglycerin Sublingual Tablets 0.4 mg", "5/31/2026", "$46.71", "Enabled"],
//   ["42", "Dental Depot", "Albuterol Sulfate (60 doses)", "5/31/2026", "$79.61", "Enabled"],
//   ["42", "Dental Depot", "Ammonia Towelette", "3/31/2027", "$14.08", "Enabled"],
//   ["42", "Dental Depot", "Adrenaline 1 mg/mL", "6/30/2026", "$31.27", "Enabled"],
//   ["42", "Dental Depot", "Adrenaline 1 mg/mL", "6/30/2026", "$31.27", "Enabled"],
//   ["42", "Dental Depot", "Naloxone HCL 0.4 mg/mL", "4/30/2026", "$43.45", "Enabled"],
// ]);

// // --- Step 6: Final Output
// const finalRevenueForecast = Array.from(forecastMap.entries()).map(([month, revenue]) => [month, revenue]);
// console.table(finalRevenueForecast);

// ─── Helpers ────────────────────────────────────────────────────────────────
function parseMonth(ym) {
  var _ym$split$map = ym.split('-').map(Number),
    _ym$split$map2 = _slicedToArray(_ym$split$map, 2),
    y = _ym$split$map2[0],
    m = _ym$split$map2[1];
  return new Date(Date.UTC(y, m - 1, 1));
}
function formatMonth(dt) {
  var y = dt.getUTCFullYear(),
    m = String(dt.getUTCMonth() + 1).padStart(2, '0');
  return "".concat(y, "-").concat(m);
}
function addDays(dt, n) {
  return new Date(dt.valueOf() + n * 864e5);
}
function addMonths(dt, n) {
  var y = dt.getUTCFullYear(),
    mo = dt.getUTCMonth() + n;
  return new Date(Date.UTC(y + Math.floor(mo / 12), mo % 12, 1));
}
function generateProjections(start, end, perMonth) {
  var result = {};
  var cur = parseMonth(start),
    last = parseMonth(end);
  while (cur <= last) {
    result[formatMonth(cur)] = perMonth;
    cur = addMonths(cur, 1);
  }
  return result;
}
function formatDate(date) {
  var year = date.getFullYear();
  var month = String(date.getMonth() + 1).padStart(2, '0'); // Months are 0-indexed
  var day = String(date.getDate()).padStart(2, '0');
  return "".concat(year, "-").concat(month);
}
function excelSerialDateToJSDate(serial) {
  var utc_days = Math.floor(serial - 25569);
  var utc_value = utc_days * 86400;
  var date = new Date(utc_value * 1000);
  return date;
}
// ─── run it ────────────────────────────────────────────────────────────────
}();
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
// Imports
var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
// Module
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\r\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\r\n\r\n<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n    <title>Protect It First Functions</title>\r\n\r\n    <!-- Office JavaScript API -->\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\r\n\r\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\r\n    <link rel=\"stylesheet\" href=\"https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css\"/>\r\n\r\n    <!-- Template styles -->\r\n    <link href=\"" + ___HTML_LOADER_IMPORT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n</head>\r\n\r\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\r\n    <header style=\"display: flex; justify-content: center; align-items: center; height: 20vh;\">\r\n        <h1 class=\"ms-font-su\">Functions</h1>\r\n    </header>\r\n    <div style=\"padding: 10px;\">\r\n        <div style=\"padding: 10px; border: 1px dashed black;\">\r\n            <h1>Revenue Predictions</h1>\r\n            <h2>Instructions</h2>\r\n            <ol>\r\n                <li>Please Don't rename the sheets </li>\r\n                <li>Keep the format of the sheets constant.</li>\r\n                <li>Drug Replenish Dates(New Kits) & Revenue Prediction contain the Drug replenishment dates and Revenue Predictions respectively</li>\r\n                <li>To Create the prediction click the button below</li>\r\n            </ol>\r\n        \r\n            <div id=\"run\" style=\"padding: 10px; background-color: cornflowerblue; color:white; cursor: pointer;text-align: center; font-weight: 700;\">\r\n                <span class=\"ms-Button-label\" style=\"text-align: center;\" >Create Revenue Predictions</span>\r\n            </div>\r\n        </div>\r\n    </div>\r\n   \r\n        <p><label id=\"item-subject\"></label></p>\r\n    </main>\r\n</body>\r\n\r\n</html>\r\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map
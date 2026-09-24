/* ============================================================
   motion.js — هەموو جوڵەکانی لاپەڕەکە
   ------------------------------------------------------------
   پێویستە motion.css لەگەڵ ئەمەدا بەکاربهێنرێت.
   لە کۆتایی <body> بانگی بکە:  <script src="js/motion.js"></script>

   جۆرەکانی جوڵە:
     .reveal        — پارچەکە بە نەرمی لە خوارەوە بەرز دەبێتەوە (یەک جار، دوایی دەوەستێت)
     .reveal.lines  — دەقەکە دێڕ بە دێڕ لە پشت پەردەوە دێتە سەرەوە
     .reveal.words  — دەقەکە وشە بە وشە دێت (بۆ ناونیشان)
     .stepper/.step — لیدی لای ڕاست: کام بەش لەبەردەستە
     پەنجەلێدان     — #boardLogo → #logoNote ، #devName → .footer.is-open
   ============================================================ */
(function(){
  "use strict";

  var reduced = window.matchMedia &&
                window.matchMedia("(prefers-reduced-motion: reduce)").matches;


  /* ---------- ناسینەوەی سکڕۆڵ لە هەموو بینەرێکدا ----------
     هەندێک بینەری HTML (وەک پێشبینینی فایل لە iOS) ڕووداوی scroll نانێرن،
     بۆیە جگە لە گوێگرتن، بە rAF شوێنی لاپەڕەکەش دەپێوین. */
  var scrollCallbacks = [];
  var lastPos = null;
  var pending = false;

  function pagePos(){
    var el = document.documentElement;
    return window.pageYOffset || el.scrollTop || document.body.scrollTop ||
           -el.getBoundingClientRect().top || 0;
  }

  function runCallbacks(){
    pending = false;
    scrollCallbacks.forEach(function(fn){ fn(); });
  }

  function requestRun(){
    if(!pending){
      pending = true;
      window.requestAnimationFrame(runCallbacks);
    }
  }

  function watchPosition(){
    var pos = pagePos();
    if(pos !== lastPos){
      lastPos = pos;
      requestRun();
    }
    window.requestAnimationFrame(watchPosition);
  }

  function onEveryScroll(fn){
    scrollCallbacks.push(fn);
  }

  window.addEventListener("scroll", requestRun, { passive: true });
  window.addEventListener("resize", requestRun, { passive: true });
  document.addEventListener("scroll", requestRun, { passive: true });
  if(window.visualViewport){
    window.visualViewport.addEventListener("scroll", requestRun, { passive: true });
    window.visualViewport.addEventListener("resize", requestRun, { passive: true });
  }
  window.requestAnimationFrame(watchPosition);

  /* دەقەکانی ".lines" دێڕ بە دێڕ لە پشت پەردەوە دێنە سەرەوە.
     دێڕەکان دوای ڕێکخستنی لاپەڕەکە دەپێوێن، بۆیە لەگەڵ هەر قەبارەیەکی شاشە دەگونجێن. */
  function splitLines(el){
    if(!el.dataset.raw) el.dataset.raw = el.innerHTML;

    // سەرەتا هەموو وشەکان جیا دەکرێنەوە بۆ پێوانەی شوێنیان
    var walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, null);
    var texts = [], t;
    while((t = walker.nextNode())) if(t.nodeValue.trim()) texts.push(t);

    texts.forEach(function(node){
      var frag = document.createDocumentFragment();
      node.nodeValue.split(/(\s+)/).forEach(function(part){
        if(!part) return;
        if(/^\s+$/.test(part)){
          frag.appendChild(document.createTextNode(" "));
        } else {
          var w = document.createElement("span");
          w.className = "lw";
          w.style.display = "inline-block";
          w.textContent = part;
          frag.appendChild(w);
        }
      });
      node.parentNode.replaceChild(frag, node);
    });

    // وشەکان بەپێی بەرزیی خۆیان دەکرێن بە دێڕ
    var nodes = [].slice.call(el.childNodes);
    var lines = [], currentTop = null, bucket = null;

    nodes.forEach(function(node){
      if(node.nodeType === 1 && node.classList.contains("lw")){
        var top = node.offsetTop;
        if(currentTop === null || Math.abs(top - currentTop) > 4){
          currentTop = top;
          bucket = [];
          lines.push(bucket);
        }
        bucket.push(node);
      } else if(bucket){
        bucket.push(node);
      }
    });

    if(!lines.length) return;

    el.innerHTML = "";
    lines.forEach(function(group, i){
      var line = document.createElement("span");
      line.className = "tline";
      var inner = document.createElement("span");
      inner.className = "tline-in";
      inner.style.transitionDelay = Math.min(i * 0.055, 0.5) + "s";
      group.forEach(function(node){
        if(node.nodeType === 1 && node.classList.contains("lw")) node.style.display = "";
        inner.appendChild(node);
      });
      line.appendChild(inner);
      el.appendChild(line);
    });
  }

  function buildLines(){
    [].forEach.call(document.querySelectorAll(".lines"), function(el){
      if(el.dataset.raw) el.innerHTML = el.dataset.raw;
      splitLines(el);
    });
  }

  var lineTimer;
  window.addEventListener("resize", function(){
    clearTimeout(lineTimer);
    lineTimer = setTimeout(function(){
      var openOnes = [].slice.call(document.querySelectorAll(".lines.is-in"));
      buildLines();
      openOnes.forEach(function(el){ el.classList.add("is-in"); });
    }, 250);
  });


  /* هەر دەقێکی ".words" دەکرێت بە وشەی جیاواز، بۆ ئەوەی یەک بە یەک بێن */
  function splitWords(el){
    var walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, null);
    var nodes = [], n;
    while((n = walker.nextNode())) if(n.nodeValue.trim()) nodes.push(n);

    nodes.forEach(function(node){
      var frag = document.createDocumentFragment();
      node.nodeValue.split(/(\s+)/).forEach(function(part){
        if(!part) return;
        if(/^\s+$/.test(part)){
          frag.appendChild(document.createTextNode(part));
        } else {
          var span = document.createElement("span");
          span.className = "word";
          span.textContent = part;
          frag.appendChild(span);
        }
      });
      node.parentNode.replaceChild(frag, node);
    });

    [].forEach.call(el.querySelectorAll(".word"), function(w, i){
      w.style.transitionDelay = Math.min(i * (reduced ? 0.02 : 0.032), 0.7) + "s";
    });
  }

  [].forEach.call(document.querySelectorAll(".words"), splitWords);

  // دێڕەکان دوای بارکردنی فۆنتەکە دەپێوین، نەک پێشی —
  // بەبێ ئەوە شکاندنی دێڕەکان دەگۆڕێت و دەق تێکدەچێت
  if(document.fonts && document.fonts.ready && document.fonts.ready.then){
    document.fonts.ready.then(buildLines);
  } else {
    buildLines();
  }


  if(!reduced && "IntersectionObserver" in window){
    document.documentElement.classList.add("js-reveal");
    var io = new IntersectionObserver(function(entries){
      entries.forEach(function(entry){
        if(entry.isIntersecting){
          entry.target.classList.add("is-in");
          io.unobserve(entry.target);
        }
      });
    }, { rootMargin: "0px 0px -6% 0px", threshold: 0.06 });

    [].forEach.call(document.querySelectorAll(".reveal"), function(el, i){
      el.style.transitionDelay = (Math.min(i % 6, 5) * 0.05) + "s";
      io.observe(el);
    });

    // پشکنینێکی زیادە لەگەڵ هەر جوڵەیەکی لاپەڕەکە
    onEveryScroll(function(){
      [].forEach.call(document.querySelectorAll(".reveal:not(.is-in)"), function(el){
        var r = el.getBoundingClientRect();
        if(r.top < window.innerHeight * 0.94 && r.bottom > 0) el.classList.add("is-in");
      });
    });
  }


  /* ---------- لیدی بەشەکان: کام بەش لەبەردەستە ---------- */
  var cards = [].slice.call(document.querySelectorAll(".card"));
  var stepper = document.getElementById("stepper");
  var steps = stepper ? [].slice.call(stepper.querySelectorAll(".step")) : [];

  if(cards.length && steps.length === cards.length){
    steps.forEach(function(step, i){
      step.addEventListener("click", function(){
        var top = cards[i].getBoundingClientRect().top + window.scrollY - 24;
        window.scrollTo({ top: top, behavior: reduced ? "auto" : "smooth" });
      });
    });

    var current = -1;

    function markActive(){
      var vh = window.innerHeight, best = -1, bestArea = 0;

      cards.forEach(function(card, i){
        var r = card.getBoundingClientRect();
        var visible = Math.min(r.bottom, vh) - Math.max(r.top, 0);
        if(visible > bestArea){ bestArea = visible; best = i; }
      });

      // لیدەکان تەنها ئەو کاتە دەردەکەون کە فۆڕمەکە لەبەردەستدایە
      var firstTop = cards[0].getBoundingClientRect().top;
      var lastBottom = cards[cards.length - 1].getBoundingClientRect().bottom;
      stepper.classList.toggle("is-visible", bestArea > 0 && firstTop < vh * 0.7 && lastBottom > 0);

      if(best !== current){
        current = best;
        steps.forEach(function(step, i){
          step.setAttribute("aria-current", i === best ? "true" : "false");
        });
      }
    }

    onEveryScroll(markActive);
    markActive();
  }


  /* ---------- کردنەوە بە پەنجەلێدان: لۆگۆ و ناوی دروستکەر ---------- */
  /* ---------- وێنەی سەرەوە و لۆگۆکە: کلیک یان دەستلێدان ---------- */
  function makeToggle(el, onToggle){
    if(!el) return;
    var timer;

    function toggle(){
      var open = onToggle();
      clearTimeout(timer);
      if(open) timer = setTimeout(function(){ if(onToggle(false) !== false){} }, 6000);
    }

    el.addEventListener("click", toggle);
    el.addEventListener("keydown", function(e){
      if(e.key === "Enter" || e.key === " "){ e.preventDefault(); toggle(); }
    });
  }

  var boardLogo = document.getElementById("boardLogo");
  var logoNote = document.getElementById("logoNote");

  makeToggle(boardLogo, function(force){
    if(!logoNote) return false;
    var open = force === false ? (logoNote.classList.remove("is-open"), false)
                               : logoNote.classList.toggle("is-open");
    logoNote.setAttribute("aria-hidden", open ? "false" : "true");
    return open;
  });

  /* ---------- ناوی دروستکەر: کلیک یان دەستلێدان ---------- */
  var footer = document.querySelector(".footer");
  var devName = document.getElementById("devName");

  if(footer && devName){
    devName.addEventListener("click", function(){
      var open = footer.classList.toggle("is-open");
      devName.setAttribute("aria-expanded", open ? "true" : "false");
      var motto = document.getElementById("devMotto");
      if(motto) motto.setAttribute("aria-hidden", open ? "false" : "true");
    });
  }

})();

// static/player.js (compatible with new App.py)
// NOTE: Bản HTML player hiện tại dùng inline script nên file này chỉ để dự phòng.
// Nếu anh Long muốn chuyển sang dùng file JS ngoài, có thể include /static/player.js.

(function () {
  const socket = io({ transports: ['websocket', 'polling'] });

  let lastEnterRoomTime = 0;
  function safeEmitEnterRoom() {
    const now = Date.now();
    if (now - lastEnterRoomTime < 1500) return;
    lastEnterRoomTime = now;
    if (roomCode && playerId && !isEnded) {
      socket.emit("player_enter_room", { room_code: roomCode, player_id: playerId });
    }
  }

  const playScreen = document.getElementById("playScreen");
  const finalScreen = document.getElementById("finalScreen");

  const elName = document.getElementById("playerName");
  const elQno = document.getElementById("playerQNo");

  const elState = document.getElementById("stateText");
  const elTimer = document.getElementById("timer");

  const elQuestionBox = document.getElementById("questionBox");
  const elOptionsBox = document.getElementById("optionsBox");

  const elFinalSub = document.getElementById("finalSub");
  const elFinalGrid = document.getElementById("finalGrid");

  const elFsTotal = document.getElementById("fs_total");
  const elFsAnswered = document.getElementById("fs_answered");
  const elFsCorrect = document.getElementById("fs_correct");
  const elFsWrong = document.getElementById("fs_wrong");
  const elFsNoAnswer = document.getElementById("fs_noanswer");
  const elFsAcc = document.getElementById("fs_acc");
  const elFsScoreWrap = document.getElementById("fs_score_wrap");
  const elFsScore = document.getElementById("fs_score");

  function qs(name){
    const p = new URLSearchParams(window.location.search);
    return (p.get(name) || "").toString();
  }

  const roomFromUrl = (qs("room") || "").trim().toUpperCase();
  const playerIdFromUrl = (qs("player_id") || "").trim();

  let roomCode = roomFromUrl || (localStorage.getItem("quiz_room_code") || "");
  let playerId = playerIdFromUrl || (localStorage.getItem("quiz_player_id") || "");

  if(roomCode) localStorage.setItem("quiz_room_code", roomCode);
  if(playerId) localStorage.setItem("quiz_player_id", playerId);

  let locked = false;
  let showQuestion = false;
  let isEnded = false;

  let startMs = null;
  let duration = 0;
  let timerHandle = null;

  function fmtPct(n){
    const x = Number(n || 0);
    return x.toFixed(2).replace(".", ",") + "%";
  }

  function setTopbar(name, qno){
    if (elName) elName.textContent = `Tên: ${name || "---"}`;
    if (elQno) elQno.textContent = `Bạn đang trả lời Câu hỏi số ${qno || "--"}`;
  }

  function setState(t){
    if (elState) elState.textContent = t || "";
  }

  function showPlay(){
    if (playScreen) playScreen.style.display = "block";
    if (finalScreen) finalScreen.style.display = "none";
  }

  function showFinal(waitingText){
    isEnded = true;
    stopTimer();
    if (elTimer) elTimer.textContent = "--";
    setLocked(true);

    if (playScreen) playScreen.style.display = "none";
    if (finalScreen) finalScreen.style.display = "block";

    if (elFinalSub) elFinalSub.textContent = waitingText || "Đang tổng hợp kết quả...";
    if (elFinalGrid) elFinalGrid.style.display = "none";
    if (elFsScoreWrap) elFsScoreWrap.style.display = "none";
  }

  function fillFinal(stats){
    if (!stats) return;
    if (elFinalSub) elFinalSub.innerHTML = `Kết quả của bạn: <b>${stats.name || ""}</b>`;
    if (elFinalGrid) elFinalGrid.style.display = "grid";

    if (elFsTotal) elFsTotal.textContent = stats.total_questions ?? 0;
    if (elFsAnswered) elFsAnswered.textContent = stats.answered ?? 0;
    if (elFsCorrect) elFsCorrect.textContent = stats.correct ?? 0;
    if (elFsWrong) elFsWrong.textContent = stats.wrong ?? 0;
    if (elFsNoAnswer) elFsNoAnswer.textContent = stats.no_answer ?? 0;
    if (elFsAcc) elFsAcc.textContent = fmtPct(stats.accuracy_pct ?? 0);

    if (elFsScoreWrap) elFsScoreWrap.style.display = "block";
    if (elFsScore) elFsScore.textContent = stats.score ?? 0;
  }

  function stopTimer(){
    if(timerHandle){
      clearInterval(timerHandle);
      timerHandle = null;
    }
  }

  function startTimer(_startMs, _duration){
    startMs = _startMs;
    duration = _duration || 0;

    stopTimer();
    timerHandle = setInterval(()=>{
      if(!startMs || !duration) return;
      const now = Date.now();
      const elapsed = Math.max(0, now - startMs);
      const remain = Math.max(0, duration*1000 - elapsed);
      const sec = Math.ceil(remain/1000);
      if (elTimer) elTimer.textContent = String(sec);
      if(remain <= 0 && elTimer) elTimer.textContent = "0";
    }, 120);
  }

  function setLocked(v){
    locked = !!v;
    const btns = document.querySelectorAll(".opt-btn");
    btns.forEach(b => b.disabled = locked || isEnded);
  }

  function renderQuestion(text){
    if(!elQuestionBox) return;
    if(showQuestion && text){
      elQuestionBox.classList.remove("hidden");
      elQuestionBox.textContent = text;
    }else{
      elQuestionBox.classList.add("hidden");
      elQuestionBox.textContent = "";
    }
  }

  function renderOptions(options){
    if(!elOptionsBox) return;
    elOptionsBox.innerHTML = "";

    if(!options || options.length !== 4){
      elOptionsBox.innerHTML = `<div class="muted">Chưa có phương án.</div>`;
      return;
    }

    options.forEach((txt, idx)=>{
      const btn = document.createElement("button");
      btn.className = "opt-btn";
      btn.type = "button";
      btn.textContent = txt;
      btn.disabled = locked || isEnded;

      btn.addEventListener("click", ()=>{
        if(locked || isEnded) return;
        btn.classList.add("selected");
        socket.emit("player_submit_answer", {
          room_code: roomCode,
          player_id: playerId,
          selected_index: idx
        });
      });

      elOptionsBox.appendChild(btn);
    });
  }

  // Connect
  if(!roomCode || !playerId){
    setTopbar("", "");
    setState("Thiếu thông tin. Vui lòng vào từ trang Join.");
    renderOptions([]);
  }else{
    safeEmitEnterRoom();
  }

  // Handlers
  socket.on("player_state", (st)=>{
    if(!st) return;

    if(isEnded || st.status === "ended"){
      showFinal("Bài thi đã kết thúc. Đang tổng hợp kết quả...");
      return;
    }

    if(st.status !== "running"){
      showPlay();
      setTopbar(st.player_name || "", "--");
      setState("Đang chờ BTC bắt đầu...");
      if (elTimer) elTimer.textContent = "--";
      renderQuestion("");
      renderOptions([]);
      setLocked(true);
      return;
    }

    isEnded = false;
    showPlay();

    showQuestion = !!st.show_question;
    setTopbar(st.player_name || "", st.q_number || (st.q_index+1));
    setState("Đang thi...");
    startTimer(st.start_ms, st.duration);
    renderQuestion(st.question_text || "");
    renderOptions(st.options || []);
    setLocked(!!st.locked);
  });

  socket.on("question_started", ()=>{
    if(!roomCode || !playerId) return;
    if(isEnded) return;
    safeEmitEnterRoom();
  });

  socket.on("time_up", ()=>{
    if(isEnded) return;
    setState("HẾT GIỜ! Chờ BTC chuyển câu...");
    setLocked(true);
  });

  socket.on("quiz_ended", ()=>{
    showFinal("Bài thi đã kết thúc. Đang tổng hợp kết quả...");
  });

  // BTC xoá người chơi -> tự động quay về trang Join
  socket.on("player_kicked", (p)=>{
    try{
      // Nếu event broadcast (room players) thì chỉ xử lý khi đúng player_id của mình
      if(p && p.player_id && playerId && String(p.player_id) !== String(playerId)){
        return;
      }
      const r = (p && p.room_code) ? String(p.room_code).toUpperCase() : roomCode;
      // Clear local storage to tránh giữ player_id cũ
      localStorage.removeItem("quiz_player_id");
      // Optional: giữ room_code để autofill
      if(r) localStorage.setItem("quiz_room_code", r);
      if(p && p.msg){
        alert(p.msg);
      }else{
        alert("Bạn đã bị BTC xoá khỏi phòng. Vui lòng tham gia lại.");
      }
      try{ socket.disconnect(); }catch(e){}
      window.location.href = `/player_join?room=${encodeURIComponent(r || "")}&kicked=1`;
    }catch(e){
      // fallback
      window.location.href = `/player_join?kicked=1`;
    }
  });

  socket.on("answer_ack", (a)=>{
    if(isEnded) return;
    if(a && a.ok){
      setLocked(true);
      setState("Đã ghi nhận đáp án.");
    }
  });

  socket.on("player_config", (cfg)=>{
    if(cfg && typeof cfg.show_question !== "undefined"){
      showQuestion = !!cfg.show_question;
      if(!isEnded){
        const txt = elQuestionBox ? elQuestionBox.textContent : "";
        renderQuestion(txt);
      }
    }
  });

  socket.on("player_final_stats", (stats)=>{
    showFinal("Đang hiển thị kết quả...");
    fillFinal(stats);
  });

})();

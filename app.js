const canvas = document.querySelector("#wheel");
const ctx = canvas.getContext("2d");
const refreshButton = document.querySelector("#refreshButton");
const pickButton = document.querySelector("#pickButton");
const spinButton = document.querySelector("#spinButton");
const goWheelButton = document.querySelector("#goWheelButton");
const repickButton = document.querySelector("#repickButton");
const nextButton = document.querySelector("#nextButton");
const countEl = document.querySelector("#count");
const updatedEl = document.querySelector("#updated");
const messageEl = document.querySelector("#message");
const screenTitle = document.querySelector("#screenTitle");
const selectionScreen = document.querySelector("#selectionScreen");
const wheelScreen = document.querySelector("#wheelScreen");
const resultScreen = document.querySelector("#resultScreen");
const nameGrid = document.querySelector("#nameGrid");
const benefitPanel = document.querySelector(".benefit-panel");
const benefitInputs = [...document.querySelectorAll(".benefit-input")];
const participantNames = document.querySelector("#participantNames");
const xlsxFallback = document.querySelector("#xlsxFallback");
const xlsxInput = document.querySelector("#xlsxInput");
const roundBadge = document.querySelector("#roundBadge");
const winnerOverlay = document.querySelector("#winnerOverlay");
const winnerRank = document.querySelector("#winnerRank");
const winnerName = document.querySelector("#winnerName");
const fireworksCanvas = document.querySelector("#fireworks");
const fireworksCtx = fireworksCanvas.getContext("2d");
const winnerList = document.querySelector("#winnerList");

const text = {
  selectionTitle: "\ud6c4\ubcf4 10\uba85 \ucd94\ub9ac\uae30",
  wheelTitle: "\uacbd\ud488 \ub8f0\ub81b",
  loading: "\uc5d1\uc140 \ud30c\uc77c\uc744 \ub2e4\uc2dc \uc77d\ub294 \uc911...",
  loadFail: "\uba85\ub2e8\uc744 \ubd88\ub7ec\uc624\uc9c0 \ubabb\ud588\uc2b5\ub2c8\ub2e4.",
  peopleUnit: "\uba85",
  updated: "\uac31\uc2e0 \uc644\ub8cc",
  loadedPrefix: "\uc5d0\uc11c \ucc38\uc11d\uc790 ",
  loadedSuffix: "\uba85\uc744 \ubd88\ub7ec\uc654\uc2b5\ub2c8\ub2e4.",
  restart: "\ucc98\uc74c\ubd80\ud130 \ub2e4\uc2dc \uc2dc\uc791\ud569\ub2c8\ub2e4.",
  scanning: "\uc2a4\ud3ec\ud2b8\ub77c\uc774\ud2b8\uac00 TOP 10 \ud6c4\ubcf4\ub97c \ucc3e\ub294 \uc911...",
  top10Done: "TOP 10 \ud6c4\ubcf4 \uc120\uc815 \uc644\ub8cc. \ub8f0\ub81b \ud654\uba74\uc73c\ub85c \uc774\ub3d9\ud569\ub2c8\ub2e4.",
  waiting: "\ucd94\ucca8 \ub300\uae30 \uc911",
  spinning: "\ub8f0\ub81b \ud68c\uc804 \uc911...",
  chooseNext: "\ub2e4\uc74c \ud68c\ucc28\ub97c \uc9c4\ud589\ud558\uac70\ub098 \ucc98\uc74c\ubd80\ud130 \ub2e4\uc2dc \uc2dc\uc791\ud560 \uc218 \uc788\uc2b5\ub2c8\ub2e4.",
  allDone: "5\uba85 \ub2f9\ucca8\uc790 \ucd94\ucca8\uc774 \uc644\ub8cc\ub418\uc5c8\uc2b5\ub2c8\ub2e4."
};

const palette = ["#123458", "#1d2a3f"];
const totalWinners = 5;
let allParticipants = [];
let candidates = [];
let wheelPeople = [];
let winners = [];
let round = 0;
let rotation = 0;
let busy = false;
let fireworksRunning = false;
let fireworksParticles = [];
let resultTimer = null;
let benefitNames = new Set();

function setMessage(value) {
  messageEl.textContent = value || "";
}

function shuffle(items) {
  const copy = [...items];
  for (let i = copy.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

function showScreen(name) {
  const selecting = name === "selection";
  const wheel = name === "wheel";
  const result = name === "result";
  selectionScreen.classList.toggle("active", selecting);
  wheelScreen.classList.toggle("active", wheel);
  resultScreen.classList.toggle("active", result);
}

function transitionToScreen(name) {
  document.body.classList.add("transitioning");
  setTimeout(() => {
    showScreen(name);
    setTimeout(() => document.body.classList.remove("transitioning"), 460);
  }, 180);
}

function setButtons(state) {
  refreshButton.disabled = busy;
  pickButton.classList.toggle("hidden", state !== "pick");
  goWheelButton.classList.toggle("hidden", state !== "readyWheel");
  repickButton.classList.toggle("hidden", state !== "spin" && state !== "afterSpin");
  spinButton.classList.toggle("hidden", state !== "spin");
  nextButton.classList.toggle("hidden", state !== "afterSpin" || round >= totalWinners);
  pickButton.disabled = busy || allParticipants.length < 10;
  spinButton.disabled = busy || wheelPeople.length < 2;
  goWheelButton.disabled = busy || candidates.length !== 10;
  repickButton.disabled = busy || allParticipants.length < 10;
}

function drawWheel() {
  const { width, height } = canvas;
  const cx = width / 2;
  const cy = height / 2;
  const radius = Math.min(width, height) * 0.46;

  ctx.clearRect(0, 0, width, height);
  ctx.save();
  ctx.translate(cx, cy);
  ctx.rotate(rotation);

  if (!wheelPeople.length) {
    ctx.beginPath();
    ctx.arc(0, 0, radius, 0, Math.PI * 2);
    ctx.fillStyle = "#111722";
    ctx.fill();
    ctx.strokeStyle = "rgba(255,255,255,.16)";
    ctx.lineWidth = 3;
    ctx.stroke();
    ctx.restore();
    return;
  }

  const segments = getWheelSegments();
  segments.forEach((segment, index) => {
    const start = segment.start - Math.PI / 2;
    const end = segment.end - Math.PI / 2;
    const person = segment.person;
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.arc(0, 0, radius, start, end);
    ctx.closePath();
    ctx.fillStyle = palette[index % palette.length];
    ctx.fill();
    ctx.strokeStyle = "rgba(255,255,255,.09)";
    ctx.lineWidth = 2;
    ctx.stroke();

    ctx.save();
    ctx.rotate(start + (end - start) / 2);
    ctx.textAlign = "right";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "#dbe9f7";
    ctx.font = "900 34px Segoe UI, sans-serif";
    ctx.fillText(person.name, radius - 34, 0, radius * 0.55);
    ctx.restore();
  });

  ctx.beginPath();
  ctx.arc(0, 0, radius, 0, Math.PI * 2);
  ctx.strokeStyle = "rgba(140,180,220,.48)";
  ctx.lineWidth = 4;
  ctx.stroke();
  ctx.restore();
}

function getWheelSegments() {
  const weights = wheelPeople.map((person) => (
    benefitNames.has(normalizeName(person.name)) ? 2 : 1
  ));
  const totalWeight = weights.reduce((sum, weight) => sum + weight, 0);
  let cursor = 0;

  return wheelPeople.map((person, index) => {
    const share = totalWeight > 0 ? weights[index] / totalWeight : 0;
    const start = cursor * Math.PI * 2;
    cursor += share;
    return {
      person,
      start,
      end: cursor * Math.PI * 2
    };
  });
}

function createNameGrid(source = allParticipants) {
  nameGrid.innerHTML = "";
  shuffle(source).forEach((person, index) => {
    const tile = document.createElement("div");
    tile.className = "name-tile";
    tile.textContent = person.name;
    tile.dataset.name = person.name;
    tile.dataset.index = index;
    nameGrid.appendChild(tile);
  });
}

function normalizeName(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function getBenefitNames() {
  return new Set(benefitInputs
    .map((input) => normalizeName(input.value))
    .filter(Boolean)
    .slice(0, 5));
}

function populateParticipantNames() {
  participantNames.innerHTML = allParticipants
    .map((person) => `<option value="${escapeHtml(person.name)}"></option>`)
    .join("");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function loadParticipants() {
  busy = true;
  setButtons("pick");
  setMessage(text.loading);
  try {
    const response = await fetch("/api/participants", { cache: "no-store" });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || text.loadFail);
    allParticipants = data.participants || [];
    populateParticipantNames();
    countEl.textContent = `${allParticipants.length}${text.peopleUnit}`;
    updatedEl.textContent = data.updatedAt || text.updated;
    resetFlow(false, false);
    setMessage(`${data.source}${text.loadedPrefix}${allParticipants.length}${text.loadedSuffix}`);
  } catch (error) {
    xlsxFallback.classList.remove("hidden");
    setMessage(error.message);
  } finally {
    busy = false;
    setButtons("pick");
  }
}

async function loadParticipantsFromFile(file) {
  if (!file) return;
  busy = true;
  setButtons("pick");
  try {
    const buffer = await file.arrayBuffer();
    allParticipants = await readParticipantsFromXlsx(buffer);
    populateParticipantNames();
    countEl.textContent = `${allParticipants.length}${text.peopleUnit}`;
    updatedEl.textContent = text.updated;
    xlsxFallback.classList.add("hidden");
    resetFlow(false, false);
  } catch (error) {
    setMessage(error.message);
  } finally {
    busy = false;
    setButtons("pick");
  }
}

async function readParticipantsFromXlsx(buffer) {
  const files = await unzipXlsx(buffer);
  const shared = readSharedStrings(files["xl/sharedStrings.xml"] || "");
  const sheet = files["xl/worksheets/sheet1.xml"];
  if (!sheet) throw new Error("sheet1.xml not found in xlsx.");
  const parser = new DOMParser();
  const xml = parser.parseFromString(sheet, "application/xml");
  const rows = [...xml.getElementsByTagName("row")];
  const people = [];

  rows.forEach((row) => {
    const rowNumber = Number(row.getAttribute("r") || 0);
    if (rowNumber < 3) return;
    const cells = {};
    [...row.getElementsByTagName("c")].forEach((cell) => {
      const ref = cell.getAttribute("r") || "";
      const col = columnNumber(ref.replace(/\d/g, ""));
      const valueNode = cell.getElementsByTagName("v")[0];
      if (!valueNode) return;
      let value = valueNode.textContent.trim();
      if (cell.getAttribute("t") === "s") {
        value = shared[Number(value)] || "";
      }
      cells[col] = value.trim();
    });
    [
      { name: cells[2], attendance: cells[3], type: "student" },
      { name: cells[8], attendance: cells[9], type: "alumni" }
    ].forEach((source) => {
      if (source.name && /^(O|o|Y|Yes|YES|1)$/.test(source.attendance || "")) {
        people.push({ name: source.name.trim(), group: source.type, row: rowNumber, type: source.type });
      }
    });
  });

  return people;
}

function columnNumber(letters) {
  return letters.toUpperCase().split("").reduce((sum, letter) => (
    sum * 26 + letter.charCodeAt(0) - 64
  ), 0);
}

function readSharedStrings(xmlText) {
  if (!xmlText) return [];
  const xml = new DOMParser().parseFromString(xmlText, "application/xml");
  return [...xml.getElementsByTagName("si")].map((si) => (
    [...si.getElementsByTagName("t")].map((node) => node.textContent).join("")
  ));
}

async function unzipXlsx(buffer) {
  const view = new DataView(buffer);
  const files = {};
  let offset = 0;
  while (offset < view.byteLength - 4) {
    if (view.getUint32(offset, true) !== 0x04034b50) {
      offset += 1;
      continue;
    }
    const method = view.getUint16(offset + 8, true);
    const compressedSize = view.getUint32(offset + 18, true);
    const fileNameLength = view.getUint16(offset + 26, true);
    const extraLength = view.getUint16(offset + 28, true);
    const nameStart = offset + 30;
    const dataStart = nameStart + fileNameLength + extraLength;
    const name = new TextDecoder().decode(new Uint8Array(buffer, nameStart, fileNameLength));
    const data = new Uint8Array(buffer, dataStart, compressedSize);
    if (method === 0) {
      files[name] = new TextDecoder().decode(data);
    } else if (method === 8) {
      files[name] = await inflateRaw(data);
    }
    offset = dataStart + compressedSize;
  }
  return files;
}

async function inflateRaw(data) {
  if (!("DecompressionStream" in window)) {
    throw new Error("이 브라우저에서는 xlsx 직접 읽기를 지원하지 않습니다. 로컬 실행 파일을 사용해주세요.");
  }
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  const buffer = await new Response(stream).arrayBuffer();
  return new TextDecoder().decode(buffer);
}

function resetFlow(keepMessage = true, animated = true) {
  candidates = [];
  wheelPeople = [];
  winners = [];
  benefitNames = getBenefitNames();
  round = 0;
  rotation = 0;
  if (resultTimer) {
    clearTimeout(resultTimer);
    resultTimer = null;
  }
  winnerOverlay.classList.add("hidden");
  stopFireworks();
  benefitPanel.classList.remove("hidden");
  if (animated) {
    transitionToScreen("selection");
  } else {
    showScreen("selection");
  }
  createNameGrid();
  drawWheel();
  setButtons("pick");
  if (keepMessage) setMessage(text.restart);
}

function pickCandidates() {
  if (busy || allParticipants.length < 10) return;
  benefitNames = getBenefitNames();
  runSpotlightSelection(allParticipants, "readyWheel");
}

function repickCandidates() {
  if (busy || allParticipants.length < 10) return;
  winnerOverlay.classList.add("hidden");
  const excluded = new Set(winners.map((winner) => normalizeName(winner.person.name)));
  const pool = allParticipants.filter((person) => !excluded.has(normalizeName(person.name)));
  benefitNames = getBenefitNames();
  transitionToScreen("selection");
  setTimeout(() => runSpotlightSelection(pool, "readyWheel"), 260);
}

function runSpotlightSelection(pool, finalState) {
  busy = true;
  setButtons("pick");
  setMessage(text.scanning);
  createNameGrid(pool);
  benefitPanel.classList.toggle("hidden", winners.length > 0);

  candidates = shuffle(pool).slice(0, 10);
  const selectedNames = new Set(candidates.map((person) => person.name));
  const tiles = [...document.querySelectorAll(".name-tile")];
  tiles.forEach((tile) => {
    tile.classList.remove("scan", "selected");
    tile.style.opacity = "1";
  });
  const active = new Set();
  const spotlightCount = Math.min(10, tiles.length);

  const clearActive = () => {
    active.forEach((tile) => tile.classList.remove("scan"));
    active.clear();
  };

  let delay = 0;
  const ticks = 22;
  for (let tick = 0; tick < ticks; tick += 1) {
    const progress = tick / (ticks - 1);
    const step = 95 + Math.pow(progress, 2.4) * 255;
    delay += step;
    setTimeout(() => {
      clearActive();
      const spotlightTiles = progress > 0.72
        ? blendTowardSelection(tiles, selectedNames, spotlightCount, progress)
        : shuffle(tiles).slice(0, spotlightCount);
      spotlightTiles.forEach((tile) => {
        tile.classList.add("scan");
        active.add(tile);
      });
    }, delay);
  }

  setTimeout(() => {
    clearActive();
    tiles.forEach((tile) => {
      const selected = selectedNames.has(tile.dataset.name);
      tile.classList.toggle("selected", selected);
      tile.style.opacity = selected ? "1" : "0.34";
    });
    setMessage(text.top10Done);
  }, delay + 420);

  setTimeout(() => {
    busy = false;
    setButtons(finalState);
  }, delay + 660);
}

function blendTowardSelection(tiles, selectedNames, count, progress) {
  const selectedTiles = tiles.filter((tile) => selectedNames.has(tile.dataset.name));
  const otherTiles = tiles.filter((tile) => !selectedNames.has(tile.dataset.name));
  const selectedCount = Math.min(count, Math.max(1, Math.round((progress - 0.72) / 0.28 * count)));
  return [
    ...shuffle(selectedTiles).slice(0, selectedCount),
    ...shuffle(otherTiles).slice(0, Math.max(0, count - selectedCount))
  ];
}

function goToWheel() {
  if (busy || candidates.length !== 10) return;
  benefitNames = getBenefitNames();
  wheelPeople = [...candidates];
  round = winners.length;
  transitionToScreen("wheel");
  setTimeout(() => prepareRound(), 220);
}

function prepareRound() {
  winnerOverlay.classList.add("hidden");
  roundBadge.textContent = `\ub2f9\ucca8! (${round + 1}/${totalWinners})`;
  setMessage(`\ub2f9\ucca8! (${round + 1}/${totalWinners}) ${text.waiting}`);
  drawWheel();
  setButtons("spin");
}

function getWinnerIndex() {
  const normalized = ((rotation % (Math.PI * 2)) + Math.PI * 2) % (Math.PI * 2);
  const pointerAngle = ((-Math.PI / 2 - normalized + Math.PI / 2) + Math.PI * 4) % (Math.PI * 2);
  const segments = getWheelSegments();
  const index = segments.findIndex((segment) => {
    return pointerAngle >= segment.start && pointerAngle < segment.end;
  });
  return index >= 0 ? index : wheelPeople.length - 1;
}

function spin() {
  if (busy || wheelPeople.length < 2 || round >= totalWinners) return;
  busy = true;
  setButtons("spin");
  setMessage(`\ub2f9\ucca8! (${round + 1}/${totalWinners}) ${text.spinning}`);

  const start = performance.now();
  const duration = 8200 + Math.random() * 900;
  const startRotation = rotation;
  const targetRotation = startRotation + Math.PI * 2 * (6.5 + Math.random() * 1.5) + Math.random() * Math.PI * 2;

  function animate(now) {
    const progress = Math.min((now - start) / duration, 1);
    const eased = 1 - Math.pow(1 - progress, 3.4);
    rotation = startRotation + (targetRotation - startRotation) * eased;
    drawWheel();

    if (progress < 1) {
      requestAnimationFrame(animate);
      return;
    }

    const winnerIndex = getWinnerIndex();
    const winner = wheelPeople[winnerIndex];
    winners.push({ person: winner });
    wheelPeople.splice(winnerIndex, 1);
    showWinner(winner);
    round += 1;
    if (round >= totalWinners) {
      setMessage(text.allDone);
      setButtons("finalPause");
      resultTimer = setTimeout(showResults, 5000);
    } else {
      busy = false;
      setButtons("afterSpin");
      setMessage(text.chooseNext);
    }
  }

  requestAnimationFrame(animate);
}

function showWinner(winner) {
  winnerRank.textContent = `\ub2f9\ucca8! (${winners.length}/${totalWinners})`;
  winnerName.textContent = winner.name;
  winnerOverlay.classList.remove("hidden");
}

function nextRound() {
  if (busy || round >= totalWinners) return;
  prepareRound();
  requestAnimationFrame(() => spin());
}

function showResults() {
  winnerList.innerHTML = winners.map((winner) => (
    `<div class="winner-card"><strong>${escapeHtml(winner.person.name)}</strong></div>`
  )).join("");
  winnerOverlay.classList.add("hidden");
  busy = false;
  transitionToScreen("result");
  setTimeout(startFireworks, 240);
  setButtons("result");
}

function resizeFireworks() {
  const rect = fireworksCanvas.getBoundingClientRect();
  const scale = window.devicePixelRatio || 1;
  fireworksCanvas.width = Math.max(1, Math.floor(rect.width * scale));
  fireworksCanvas.height = Math.max(1, Math.floor(rect.height * scale));
  fireworksCtx.setTransform(scale, 0, 0, scale, 0, 0);
}

function launchFirework() {
  const rect = fireworksCanvas.getBoundingClientRect();
  const x = rect.width * (0.12 + Math.random() * 0.76);
  const y = rect.height * (0.12 + Math.random() * 0.42);
  const colors = ["#d7ebff", "#86b4df", "#f8fbff", "#9cc8ef"];
  const count = 42;
  for (let i = 0; i < count; i += 1) {
    const angle = (Math.PI * 2 * i) / count + Math.random() * 0.08;
    const speed = 1.8 + Math.random() * 4.2;
    fireworksParticles.push({
      x,
      y,
      vx: Math.cos(angle) * speed,
      vy: Math.sin(angle) * speed,
      life: 68 + Math.random() * 26,
      age: 0,
      color: colors[Math.floor(Math.random() * colors.length)]
    });
  }
}

function drawFireworks() {
  if (!fireworksRunning) return;
  const rect = fireworksCanvas.getBoundingClientRect();
  fireworksCtx.clearRect(0, 0, rect.width, rect.height);
  fireworksCtx.globalCompositeOperation = "lighter";

  fireworksParticles = fireworksParticles.filter((particle) => particle.age < particle.life);
  fireworksParticles.forEach((particle) => {
    particle.age += 1;
    particle.x += particle.vx;
    particle.y += particle.vy;
    particle.vy += 0.035;
    particle.vx *= 0.992;
    particle.vy *= 0.992;
    const alpha = 1 - particle.age / particle.life;
    fireworksCtx.globalAlpha = alpha;
    fireworksCtx.fillStyle = particle.color;
    fireworksCtx.beginPath();
    fireworksCtx.arc(particle.x, particle.y, 2.2, 0, Math.PI * 2);
    fireworksCtx.fill();
  });

  fireworksCtx.globalAlpha = 1;
  requestAnimationFrame(drawFireworks);
}

function startFireworks() {
  resizeFireworks();
  fireworksParticles = [];
  fireworksRunning = true;
  launchFirework();
  drawFireworks();
}

function stopFireworks() {
  fireworksRunning = false;
  fireworksParticles = [];
  if (fireworksCtx) {
    const rect = fireworksCanvas.getBoundingClientRect();
    fireworksCtx.clearRect(0, 0, rect.width, rect.height);
  }
}

setInterval(() => {
  if (fireworksRunning) launchFirework();
}, 820);

window.addEventListener("resize", () => {
  if (fireworksRunning) resizeFireworks();
});

refreshButton.addEventListener("click", loadParticipants);
xlsxInput.addEventListener("change", () => loadParticipantsFromFile(xlsxInput.files[0]));
pickButton.addEventListener("click", pickCandidates);
goWheelButton.addEventListener("click", goToWheel);
repickButton.addEventListener("click", repickCandidates);
spinButton.addEventListener("click", spin);
nextButton.addEventListener("click", nextRound);
winnerOverlay.addEventListener("click", () => {
  if (!busy) winnerOverlay.classList.add("hidden");
});

drawWheel();
loadParticipants();

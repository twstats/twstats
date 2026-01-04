import "./App.css";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import * as XLSX from "xlsx";

type TagGroup =
  | "Scoring"
  | "Set piece & breakdown"
  | "Attack"
  | "Defence & discipline"
  | "Restarts"
  | "Game Flow";

type TagColor = "green" | "blue" | "orange" | "purple" | "yellow" | "red";

type TagDef = {
  id: string;
  label: string;
  key: string; // hotkey
  group: TagGroup;
  color: TagColor;
  playerRequired?: boolean; // false = logs without player
};

type PlayerStatus = "on" | "bench" | "off";

type Player = {
  id: string;
  num: number;
  name: string;
  status: PlayerStatus;
};

type EventLog = {
  id: string;
  tsSec: number; // match clock (synced)
  videoSec: number;
  playerId: string | null;
  tagId: string;
  meta?: Record<string, any>;
};

const STORAGE_KEY = "twstats_state_v5";
const VIDEO_H_KEY = "twstats_video_height_v1";

const TAGS: TagDef[] = [
  // Scoring
  { id: "try", label: "Try", key: "1", group: "Scoring", color: "green" },
  { id: "conv_made", label: "Conv made", key: "2", group: "Scoring", color: "green" },
  { id: "conv_miss", label: "Conv miss", key: "3", group: "Scoring", color: "green" },
  { id: "pen_goal", label: "Penalty goal", key: "4", group: "Scoring", color: "green" },
  { id: "pen_miss", label: "Penalty miss", key: "5", group: "Scoring", color: "green" },
  { id: "drop_goal", label: "Drop goal", key: "6", group: "Scoring", color: "green" },
  { id: "pen_try", label: "Penalty try", key: "7", group: "Scoring", color: "green" },

  // Set piece & breakdown
  { id: "scrum_won", label: "Scrum won", key: "q", group: "Set piece & breakdown", color: "blue" },
  { id: "scrum_lost", label: "Scrum lost", key: "w", group: "Set piece & breakdown", color: "blue" },
  { id: "lineout_won", label: "Lineout won", key: "e", group: "Set piece & breakdown", color: "blue" },
  { id: "lineout_lost", label: "Lineout lost", key: "r", group: "Set piece & breakdown", color: "blue" },
  { id: "ruck_won", label: "Ruck won", key: "t", group: "Set piece & breakdown", color: "blue" },
  { id: "ruck_lost", label: "Ruck lost", key: "y", group: "Set piece & breakdown", color: "blue" },
  { id: "maul_won", label: "Maul won", key: "u", group: "Set piece & breakdown", color: "blue" },
  { id: "maul_lost", label: "Maul lost", key: "i", group: "Set piece & breakdown", color: "blue" },

  // Attack
  { id: "carry", label: "Carry", key: "a", group: "Attack", color: "orange" },
  { id: "pass", label: "Pass", key: "s", group: "Attack", color: "orange" },
  { id: "offload", label: "Offload", key: "d", group: "Attack", color: "orange" },
  { id: "linebreak", label: "Linebreak", key: "f", group: "Attack", color: "orange" },
  { id: "lb_assist", label: "LB assist", key: "g", group: "Attack", color: "orange" },
  { id: "kick_in_play", label: "Kick in play", key: "z", group: "Attack", color: "orange" },
  { id: "box_kick", label: "Box kick", key: "x", group: "Attack", color: "orange" },
  { id: "grubber", label: "Grubber", key: "c", group: "Attack", color: "orange" },
  { id: "chip_kick", label: "Chip kick", key: "v", group: "Attack", color: "orange" },
  { id: "kick_touch", label: "Kick to touch", key: "b", group: "Attack", color: "orange" },
  { id: "catch", label: "Catch", key: "n", group: "Attack", color: "orange" },
  { id: "def_beaten", label: "Defender beaten", key: "m", group: "Attack", color: "orange" },

  // Defence & discipline
  { id: "tackle", label: "Tackle", key: "j", group: "Defence & discipline", color: "purple" },
  { id: "dom_tackle", label: "Dom tackle", key: "k", group: "Defence & discipline", color: "purple" },
  { id: "missed_tackle", label: "Missed tackle", key: "l", group: "Defence & discipline", color: "purple" },
  { id: "turnover_won", label: "Turnover won", key: "o", group: "Defence & discipline", color: "purple" },
  { id: "turnover_conc", label: "Turnover conceded", key: "p", group: "Defence & discipline", color: "purple" },
  { id: "jackal_pen", label: "Jackal pen won", key: ";", group: "Defence & discipline", color: "purple" },
  { id: "pen_conc", label: "Penalty conceded", key: "h", group: "Defence & discipline", color: "purple" },
  { id: "knock_on", label: "Knock on", key: ".", group: "Defence & discipline", color: "purple" },
  { id: "handling", label: "Handling error", key: "/", group: "Defence & discipline", color: "purple" },
  { id: "yellow", label: "Yellow card", key: "9", group: "Defence & discipline", color: "yellow" },
  { id: "red", label: "Red card", key: "8", group: "Defence & discipline", color: "red" },

  // Restarts
  { id: "drop_start", label: "Drop Start", key: "d", group: "Restarts", color: "purple", playerRequired: false },

  // Game Flow
  { id: "end_play", label: "End of Play", key: "e", group: "Game Flow", color: "orange", playerRequired: false },
  { id: "sub", label: "Substitution", key: "s", group: "Game Flow", color: "blue", playerRequired: false },
];

function uid() {
  return (crypto as any)?.randomUUID
    ? (crypto as any).randomUUID()
    : Math.random().toString(16).slice(2) + Date.now().toString(16);
}

function secToClock(sec: number) {
  const total = Math.max(0, Math.floor(sec));
  const m = Math.floor(total / 60);
  const s = total % 60;
  return `${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`;
}

export default function App() {
  // players: 1-15 ON, 16-23 BENCH
  const [players, setPlayers] = useState<Player[]>(
    Array.from({ length: 23 }, (_, i) => {
      const num = i + 1;
      return { id: `p${num}`, num, name: `Player ${num}`, status: num <= 15 ? "on" : "bench" };
    })
  );

  const videoRef = useRef<HTMLVideoElement | null>(null);
  const rafRef = useRef<number | null>(null);
  const resizingRef = useRef(false);

  const [selectedPlayerId, setSelectedPlayerId] = useState<string>("p1");
  const [events, setEvents] = useState<EventLog[]>([]);
  const [urlInput, setUrlInput] = useState("");

  // match clock is synced: match = videoTime + offset
  const [matchClockOffset, setMatchClockOffset] = useState<number>(0);
  const [matchTimeSec, setMatchTimeSec] = useState<number>(0);

  // video extras
  const [playbackRate, setPlaybackRate] = useState<number>(1);
  const [skipBack, setSkipBack] = useState<number>(2);
  const [skipFwd, setSkipFwd] = useState<number>(5);

  // analytics pages
  const [page, setPage] = useState<"raw" | "totals" | "players">("raw");

  // subs controls
  const [subOffId, setSubOffId] = useState<string>("p1");
  const [subOnId, setSubOnId] = useState<string>("p16");

  // editable names
  const [editingPlayerId, setEditingPlayerId] = useState<string | null>(null);

  // resizable video height
  const [videoHeight, setVideoHeight] = useState<number>(() => {
    const raw = localStorage.getItem(VIDEO_H_KEY);
    const n = raw ? Number(raw) : 460;
    if (!Number.isFinite(n)) return 460;
    return Math.min(1000, Math.max(260, n));
  });

  // --- maps
  const tagMap = useMemo(() => new Map(TAGS.map((t) => [t.id, t])), []);
  const playerMap = useMemo(() => new Map(players.map((p) => [p.id, p])), [players]);

  const onFieldPlayers = useMemo(() => players.filter((p) => p.status === "on"), [players]);
  const benchCandidates = useMemo(() => players.filter((p) => p.status === "bench" || p.status === "off"), [players]);

  // keep selected player valid (must be ON)
  useEffect(() => {
    const p = playerMap.get(selectedPlayerId);
    if (p && p.status === "on") return;
    if (onFieldPlayers.length > 0) setSelectedPlayerId(onFieldPlayers[0].id);
  }, [selectedPlayerId, playerMap, onFieldPlayers]);

  // --- persist height
  useEffect(() => {
    localStorage.setItem(VIDEO_H_KEY, String(videoHeight));
  }, [videoHeight]);

  // --- drag resize handlers
  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      if (!resizingRef.current) return;
      setVideoHeight((h) => Math.min(1000, Math.max(260, h + e.movementY)));
    };
    const onUp = () => (resizingRef.current = false);

    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, []);


  const resetVideoSize = () => setVideoHeight(460);

  // --- load/save full state
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw) as Partial<{
        players: Player[];
        selectedPlayerId: string;
        events: EventLog[];
        matchClockOffset: number;
        playbackRate: number;
        skipBack: number;
        skipFwd: number;
      }>;

      if (Array.isArray(parsed.players)) setPlayers(parsed.players);
      if (typeof parsed.selectedPlayerId === "string") setSelectedPlayerId(parsed.selectedPlayerId);
      if (Array.isArray(parsed.events)) setEvents(parsed.events);
      if (typeof parsed.matchClockOffset === "number") setMatchClockOffset(parsed.matchClockOffset);

      if (typeof parsed.playbackRate === "number") setPlaybackRate(parsed.playbackRate);
      if (typeof parsed.skipBack === "number") setSkipBack(parsed.skipBack);
      if (typeof parsed.skipFwd === "number") setSkipFwd(parsed.skipFwd);
    } catch {}
  }, []);

  useEffect(() => {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({ players, selectedPlayerId, events, matchClockOffset, playbackRate, skipBack, skipFwd })
    );
  }, [players, selectedPlayerId, events, matchClockOffset, playbackRate, skipBack, skipFwd]);

  // --- sync match clock from video
  const syncClockFromVideo = useCallback(() => {
    const v = videoRef.current;
    if (!v) return;
    const mt = (v.currentTime || 0) + matchClockOffset;
    setMatchTimeSec(Math.max(0, mt));
    rafRef.current = requestAnimationFrame(syncClockFromVideo);
  }, [matchClockOffset]);

  useEffect(() => {
    const v = videoRef.current;
    if (!v) return;

    const onPlay = () => {
      if (rafRef.current) cancelAnimationFrame(rafRef.current);
      rafRef.current = requestAnimationFrame(syncClockFromVideo);
    };

    const onPause = () => {
      if (rafRef.current) cancelAnimationFrame(rafRef.current);
      rafRef.current = null;
      const mt = (v.currentTime || 0) + matchClockOffset;
      setMatchTimeSec(Math.max(0, mt));
    };

    const onEnded = () => onPause();

    v.addEventListener("play", onPlay);
    v.addEventListener("pause", onPause);
    v.addEventListener("ended", onEnded);

    return () => {
      v.removeEventListener("play", onPlay);
      v.removeEventListener("pause", onPause);
      v.removeEventListener("ended", onEnded);
      if (rafRef.current) cancelAnimationFrame(rafRef.current);
    };
  }, [syncClockFromVideo, matchClockOffset]);

  // playback rate
  useEffect(() => {
    const v = videoRef.current;
    if (!v) return;
    v.playbackRate = playbackRate;
  }, [playbackRate]);

  const playPause = () => {
    const v = videoRef.current;
    if (!v) return;
    if (v.paused) v.play().catch(() => {});
    else v.pause();
  };

  const resetMatchClockToZeroHere = () => {
    const v = videoRef.current;
    if (!v) return;
    v.pause();
    setMatchClockOffset(-(v.currentTime || 0));
    setMatchTimeSec(0);
  };

  const jump = (delta: number) => {
    const v = videoRef.current;
    if (!v) return;
    const dur = Number.isFinite(v.duration) ? v.duration : Infinity;
    v.currentTime = Math.max(0, Math.min(dur, (v.currentTime || 0) + delta));
  };

  const addEvent = (tagId: string, opts?: { playerId?: string | null; meta?: Record<string, any> }) => {
    const tag = tagMap.get(tagId);
    if (!tag) return;

    const requiresPlayer = tag.playerRequired !== false;
    const playerId = requiresPlayer ? (opts?.playerId ?? selectedPlayerId) : (opts?.playerId ?? null);

    if (requiresPlayer) {
      const p = playerId ? playerMap.get(playerId) : null;
      if (!p || p.status !== "on") return;
    }

    const v = videoRef.current;
    const videoSec = v ? v.currentTime || 0 : 0;

    const ev: EventLog = {
      id: uid(),
      tsSec: matchTimeSec,
      videoSec,
      playerId,
      tagId,
      meta: opts?.meta,
    };

    setEvents((prev) => [ev, ...prev]);
  };

  const undoLast = () => setEvents((prev) => prev.slice(1));
  const deleteEvent = (id: string) => setEvents((prev) => prev.filter((e) => e.id !== id));

  // substitutions
  const doSubstitution = () => {
    const off = playerMap.get(subOffId);
    const on = playerMap.get(subOnId);
    if (!off || !on) return;
    if (off.status !== "on") return;
    if (on.status === "on") return;

    setPlayers((prev) =>
      prev.map((p) => {
        if (p.id === off.id) return { ...p, status: "off" };
        if (p.id === on.id) return { ...p, status: "on" };
        return p;
      })
    );

    addEvent("sub", {
      playerId: null,
      meta: { offId: off.id, onId: on.id, offNum: off.num, onNum: on.num },
    });

    setSelectedPlayerId(on.id);
  };

  // editable names
  const updatePlayerName = (playerId: string, name: string) => {
    setPlayers((prev) => prev.map((p) => (p.id === playerId ? { ...p, name } : p)));
  };

  // hotkeys
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      const tn = target?.tagName?.toLowerCase();
      const isTyping = tn === "input" || tn === "textarea" || (target as any)?.isContentEditable;
      if (isTyping) return;

      if (e.code === "Space") {
        e.preventDefault();
        playPause();
        return;
      }

      if (e.key.toLowerCase() === "u") {
        e.preventDefault();
        undoLast();
        return;
      }

      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") {
        e.preventDefault();
        undoLast();
        return;
      }

      const hit = TAGS.find((t) => t.key.toLowerCase() === e.key.toLowerCase());
      if (hit) {
        e.preventDefault();
        addEvent(hit.id);
      }
    };

    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [matchTimeSec, selectedPlayerId, players]);

  function handleVideoFile(file: File | null) {
    const v = videoRef.current;
    if (!v || !file) return;
    const objectUrl = URL.createObjectURL(file);
    v.src = objectUrl;
    v.load();
    v.pause();
  }

  function loadFromUrl() {
    const v = videoRef.current;
    if (!v) return;
    if (!urlInput.trim()) return;
    v.src = urlInput.trim();
    v.load();
    v.pause();
  }

  // groups
  const groups = useMemo(() => {
    const pick = (g: TagGroup) => TAGS.filter((t) => t.group === g);
    return {
      Scoring: pick("Scoring"),
      "Set piece & breakdown": pick("Set piece & breakdown"),
      Attack: pick("Attack"),
      "Defence & discipline": pick("Defence & discipline"),
      Restarts: pick("Restarts"),
      "Game Flow": pick("Game Flow"),
    } as const;
  }, []);

  // analytics
  const totalsByTag = useMemo(() => {
    const map = new Map<string, number>();
    for (const e of events) map.set(e.tagId, (map.get(e.tagId) || 0) + 1);
    return Array.from(map.entries())
      .map(([tagId, total]) => ({ tagId, total, tag: tagMap.get(tagId) }))
      .sort((a, b) => b.total - a.total);
  }, [events, tagMap]);

  const allTagIdsSeen = useMemo(() => {
    const s = new Set<string>();
    for (const e of events) s.add(e.tagId);
    return Array.from(s.values()).sort((a, b) => (tagMap.get(a)?.label || a).localeCompare(tagMap.get(b)?.label || b));
  }, [events, tagMap]);

  const statsByPlayer = useMemo(() => {
    const map = new Map<string, Map<string, number>>();
    for (const e of events) {
      if (!e.playerId) continue;
      if (!map.has(e.playerId)) map.set(e.playerId, new Map());
      const inner = map.get(e.playerId)!;
      inner.set(e.tagId, (inner.get(e.tagId) || 0) + 1);
    }

    const rows = players.map((p) => {
      const inner = map.get(p.id) || new Map<string, number>();
      const counts: Record<string, number> = {};
      for (const tagId of allTagIdsSeen) counts[tagId] = inner.get(tagId) || 0;
      const total = Object.values(counts).reduce((a, b) => a + b, 0);
      return { player: p, counts, total };
    });

    return rows.sort((a, b) => b.total - a.total);
  }, [events, players, allTagIdsSeen]);

  // export xlsx (3 sheets)
  const exportXlsx = () => {
    const raw = events
      .slice()
      .reverse()
      .map((e) => {
        const t = tagMap.get(e.tagId);
        const p = e.playerId ? playerMap.get(e.playerId) : null;
        return {
          MatchTime: secToClock(e.tsSec),
          MatchTimeSec: Number(e.tsSec.toFixed(2)),
          VideoTime: secToClock(e.videoSec),
          VideoTimeSec: Number(e.videoSec.toFixed(2)),
          Tag: t?.label ?? e.tagId,
          TagGroup: t?.group ?? "",
          Hotkey: t?.key ?? "",
          PlayerNum: p?.num ?? "",
          Player: p ? p.name : "",
          Meta: e.meta ? JSON.stringify(e.meta) : "",
        };
      });

    const totals = totalsByTag.map((r) => ({
      Tag: r.tag?.label ?? r.tagId,
      TagId: r.tagId,
      Group: r.tag?.group ?? "",
      Total: r.total,
    }));

    const playerWide = statsByPlayer.map((row) => {
      const out: Record<string, any> = {
        PlayerNum: row.player.num,
        Player: row.player.name,
        Status: row.player.status,
        Total: row.total,
      };
      for (const tagId of allTagIdsSeen) {
        const label = tagMap.get(tagId)?.label ?? tagId;
        out[label] = row.counts[tagId] || 0;
      }
      return out;
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(raw), "Raw Events");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(totals), "Totals");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(playerWide), "Player Stats");

    XLSX.writeFile(wb, `TWstats_export_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div className="app">
      <div className="topbar">
        <div className="title">TWstats</div>
        <div className="hintkeys">
          Space = play/pause • U / Ctrl+Z = undo • End of Play = (hotkey: {tagMap.get("end_play")?.key?.toUpperCase()}) • Drop Start = (hotkey:{" "}
          {tagMap.get("drop_start")?.key?.toUpperCase()}) • Drag under video to resize (dbl-click to reset)
        </div>
      </div>

      <div className="main">
        {/* LEFT: PLAYERS */}
        <div className="card">
          <div className="cardInner">
            <div className="sectionTitle">PLAYERS (On-field only)</div>

            <div className="playerList">
              {players.map((p) => {
                const selectable = p.status === "on";
                const active = p.id === selectedPlayerId;
                const isEditing = editingPlayerId === p.id;

                return (
                  <div
                    key={p.id}
                    className={`playerRow ${active ? "active" : ""} ${selectable ? "" : "disabled"}`}
                    onClick={() => selectable && setSelectedPlayerId(p.id)}
                    title={
                      selectable
                        ? "On field"
                        : p.status === "bench"
                        ? "Bench (not selectable until subbed on)"
                        : "Off field"
                    }
                    style={{ opacity: selectable ? 1 : 0.45, cursor: selectable ? "pointer" : "not-allowed" }}
                  >
                    <div className="playerNum">#{p.num}</div>

                    <div className="playerName" style={{ flex: 1 }}>
                      {isEditing ? (
                        <input
                          className="input"
                          style={{ minWidth: 0, width: "100%", height: 32 }}
                          value={p.name}
                          autoFocus
                          onClick={(e) => e.stopPropagation()}
                          onChange={(e) => updatePlayerName(p.id, e.target.value)}
                          onBlur={() => setEditingPlayerId(null)}
                          onKeyDown={(e) => {
                            if (e.key === "Enter" || e.key === "Escape") setEditingPlayerId(null);
                          }}
                        />
                      ) : (
                        <span
                          onClick={(e) => {
                            e.stopPropagation();
                            setEditingPlayerId(p.id);
                          }}
                          title="Click to edit name"
                          style={{ display: "inline-block" }}
                        >
                          {p.name}
                        </span>
                      )}
                    </div>

                    <span className="badge" style={{ marginLeft: "auto" }}>
                      {p.status.toUpperCase()}
                    </span>

                    <button
                      className="btn"
                      style={{ height: 28, padding: "0 10px", marginLeft: 10 }}
                      onClick={(e) => {
                        e.stopPropagation();
                        setEditingPlayerId(p.id);
                      }}
                      title="Edit name"
                    >
                      Edit
                    </button>
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        {/* MIDDLE: TAGS */}
        <div className="card">
          <div className="cardInner">
            {(Object.keys(groups) as Array<keyof typeof groups>).map((gname) => (
              <div className="group" key={gname}>
                <div className="groupName">{gname}</div>
                <div className="pills">
                  {groups[gname].map((t) => (
                    <button
                      key={t.id}
                      className={`pill ${t.color}`}
                      onClick={() => addEvent(t.id)}
                      title={`Hotkey: ${t.key}`}
                    >
                      {t.label}
                      <span className="keyBubble">{t.key.toUpperCase()}</span>
                    </button>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* RIGHT: VIDEO + CLOCK/ANALYTICS */}
        <div className="right">
          {/* VIDEO */}
          <div className="card">
            <div className="cardInner">
              <div className="videoHeaderRow">
                <div className="sectionTitle">MATCH VIDEO</div>

                <div className="clockBtns">
                  <button className="btn" onClick={() => jump(-skipBack)} title="Skip back">
                    -{skipBack}s
                  </button>
                  <button className="btn" onClick={() => jump(skipFwd)} title="Skip forward">
                    +{skipFwd}s
                  </button>

                  <button className="btn" onClick={() => setPlaybackRate(0.5)}>
                    0.5x
                  </button>
                  <button className="btn" onClick={() => setPlaybackRate(0.75)}>
                    0.75x
                  </button>
                  <button className="btn green" onClick={() => setPlaybackRate(1)}>
                    1x
                  </button>
                  <button className="btn" onClick={() => setPlaybackRate(1.25)}>
                    1.25x
                  </button>
                </div>
              </div>

              <div className="fileRow" style={{ marginBottom: 10 }}>
                <span>File:</span>
                <input type="file" accept="video/*" onChange={(e) => handleVideoFile(e.target.files?.[0] ?? null)} />
                <span>or URL:</span>

                <div className="urlRow">
                  <input
                    className="input"
                    value={urlInput}
                    onChange={(e) => setUrlInput(e.target.value)}
                    placeholder="https://example.com/video.mp4"
                  />
                  <button className="btn green" onClick={loadFromUrl}>
                    LOAD
                  </button>
                </div>
              </div>

              <div className="fileRow" style={{ marginBottom: 10 }}>
                <span>Skip back:</span>
                <input
                  className="input"
                  style={{ minWidth: 90, width: 90 }}
                  value={skipBack}
                  onChange={(e) => setSkipBack(Math.max(0, Number(e.target.value) || 0))}
                />
                <span>Skip fwd:</span>
                <input
                  className="input"
                  style={{ minWidth: 90, width: 90 }}
                  value={skipFwd}
                  onChange={(e) => setSkipFwd(Math.max(0, Number(e.target.value) || 0))}
                />
                <span>Rate:</span>
                <span className="badge">{playbackRate}x</span>
              </div>

              <video ref={videoRef} className="video" controls style={{ height: `${videoHeight}px` }} />

              <div
                className="resizeHandle"
                onMouseDown={() => (resizingRef.current = true)}
                onDoubleClick={resetVideoSize}
                title="Drag to resize. Double-click to reset."
              />
            </div>
          </div>

          {/* LOWER */}
          <div className="lower">
            {/* CLOCK + SUBS */}
            <div className="card">
              <div className="cardInner">
                <div className="sectionTitle">MATCH CLOCK (Synced to video)</div>

                <div className="clockBox">
                  <div>
                    <div className="clockBig">{secToClock(matchTimeSec)}</div>
                    <div className="hintkeys" style={{ marginTop: 6 }}>
                      Clock = video time (+ offset). RESET CLOCK sets match clock to 00:00 at current video position.
                    </div>
                  </div>

                  <div className="clockBtns">
                    <button className="btn green" onClick={playPause}>
                      PLAY / PAUSE
                    </button>
                    <button className="btn" onClick={resetMatchClockToZeroHere}>
                      RESET CLOCK
                    </button>
                    <button className="btn" onClick={() => addEvent("end_play")} title="End of Play">
                      END OF PLAY
                    </button>
                    <button className="btn" onClick={() => addEvent("drop_start")} title="Drop Start">
                      DROP START
                    </button>
                    <button className="btn" onClick={undoLast} title="Undo last event">
                      UNDO
                    </button>
                    <button className="btn" onClick={exportXlsx} title="Export Excel (3 sheets)">
                      EXPORT XLSX
                    </button>
                  </div>
                </div>

                <div style={{ marginTop: 14 }}>
                  <div className="sectionTitle">SUBSTITUTIONS</div>
                  <div className="fileRow">
                    <span>Off:</span>
                    <select className="btn" value={subOffId} onChange={(e) => setSubOffId(e.target.value)} style={{ height: 36 }}>
                      {onFieldPlayers.map((p) => (
                        <option key={p.id} value={p.id}>
                          #{p.num} {p.name}
                        </option>
                      ))}
                    </select>

                    <span>On:</span>
                    <select className="btn" value={subOnId} onChange={(e) => setSubOnId(e.target.value)} style={{ height: 36 }}>
                      {benchCandidates.map((p) => (
                        <option key={p.id} value={p.id}>
                          #{p.num} {p.name} ({p.status})
                        </option>
                      ))}
                    </select>

                    <button className="btn green" onClick={doSubstitution} title="Logs SUB + makes player selectable">
                      CONFIRM SUB
                    </button>
                  </div>
                </div>
              </div>
            </div>

            {/* ANALYTICS */}
            <div className="card">
              <div className="cardInner">
                <div className="videoHeaderRow">
                  <div className="sectionTitle">ANALYTICS</div>
                  <div className="clockBtns">
                    <button className={`btn ${page === "raw" ? "green" : ""}`} onClick={() => setPage("raw")}>
                      RAW
                    </button>
                    <button className={`btn ${page === "totals" ? "green" : ""}`} onClick={() => setPage("totals")}>
                      TOTALS
                    </button>
                    <button className={`btn ${page === "players" ? "green" : ""}`} onClick={() => setPage("players")}>
                      PLAYERS
                    </button>
                  </div>
                </div>

                {events.length === 0 ? (
                  <div className="feedEmpty">No events yet. Play the video and start tagging.</div>
                ) : page === "raw" ? (
                  <div className="feed">
                    {events.map((ev) => {
                      const t = tagMap.get(ev.tagId);
                      const p = ev.playerId ? playerMap.get(ev.playerId) : null;

                      return (
                        <div key={ev.id} className="feedItem">
                          <div className="feedTop">
                            <span>{secToClock(ev.tsSec)}</span>
                            <span>{p ? `#${p.num} ${p.name}` : "—"}</span>
                            <span className="badge">{t?.group ?? "Event"}</span>
                          </div>

                          <div className="feedMain">
                            <span className="feedLabel">{t?.label ?? ev.tagId}</span>
                            <span className="feedMeta">
                              {ev.videoSec.toFixed(2)}s{" "}
                              <button
                                className="btn"
                                style={{ height: 28, padding: "0 10px" }}
                                onClick={() => deleteEvent(ev.id)}
                                title="Delete this row"
                              >
                                Delete
                              </button>
                            </span>
                          </div>

                          {ev.tagId === "sub" && ev.meta ? (
                            <div className="hintkeys" style={{ marginTop: 8 }}>
                              Sub: #{playerMap.get(ev.meta.offId)?.num} off → #{playerMap.get(ev.meta.onId)?.num} on
                            </div>
                          ) : null}
                        </div>
                      );
                    })}
                  </div>
                ) : page === "totals" ? (
                  <div className="feed">
                    {totalsByTag.map((r) => (
                      <div key={r.tagId} className="feedItem">
                        <div className="feedTop">
                          <span className="badge">{r.tag?.group ?? ""}</span>
                          <span className="badge">Total</span>
                        </div>
                        <div className="feedMain">
                          <span className="feedLabel">{r.tag?.label ?? r.tagId}</span>
                          <span className="feedMeta">{r.total}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="feed">
                    {statsByPlayer.map((row) => (
                      <div
                        key={row.player.id}
                        className="feedItem"
                        style={{ opacity: row.total > 0 || row.player.status === "on" ? 1 : 0.7 }}
                      >
                        <div className="feedTop">
                          <span>
                            #{row.player.num} {row.player.name}
                          </span>
                          <span className="badge">{row.player.status.toUpperCase()}</span>
                          <span className="badge">Total: {row.total}</span>
                        </div>

                        <div className="feedMain" style={{ flexWrap: "wrap" as any }}>
                          {allTagIdsSeen
                            .map((tagId) => ({ tagId, label: tagMap.get(tagId)?.label ?? tagId, n: row.counts[tagId] || 0 }))
                            .filter((x) => x.n > 0)
                            .map((x) => (
                              <span key={x.tagId} className="badge" style={{ marginRight: 8, marginTop: 8 }}>
                                {x.label}: {x.n}
                              </span>
                            ))}
                          {row.total === 0 ? <span className="feedEmpty">No stats yet</span> : null}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>

      <div style={{ padding: "0 14px 14px", color: "rgba(255,255,255,.65)", fontWeight: 700, fontSize: 12 }}>
        Click a player name to edit it. Only <b>ON</b> players are selectable for tagging. Bench/off players become selectable only after a substitution.
      </div>
    </div>
  );
}

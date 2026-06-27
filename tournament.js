(function () {
  'use strict';

  var workbookPromise;
  var tournamentData;
  var selectedGroup = 'all';
  var liveTimer;
  var TOURNAMENT_DATE = '2026-07-04';

  function esc(value) {
    return String(value == null ? '' : value).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  }

  function loadWorkbook() {
    if (!workbookPromise) {
      var embeddedWorkbook = function () {
        if (!window.TOURNAMENT_WORKBOOK_BASE64) {
          throw new Error('Dati incorporati del torneo non disponibili.');
        }
        var binary = atob(window.TOURNAMENT_WORKBOOK_BASE64);
        var bytes = new Uint8Array(binary.length);
        for (var i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        return XLSX.read(bytes, { type: 'array' });
      };

      if (window.location.protocol === 'file:') {
        workbookPromise = Promise.resolve().then(embeddedWorkbook);
      } else {
        workbookPromise = fetch('./12HBosaro_manager.xlsx?v=' + Date.now(), { cache: 'no-store' })
          .then(function (response) {
            if (!response.ok) throw new Error('Workbook manager non disponibile (' + response.status + ')');
            return response.arrayBuffer();
          })
          .then(function (buffer) { return XLSX.read(buffer, { type: 'array' }); })
          .catch(function (error) {
            console.warn('Uso i dati incorporati del torneo:', error);
            return embeddedWorkbook();
          });
      }
    }
    return workbookPromise;
  }

  function recordsBetween(rows, marker, nextMarker) {
    var start = rows.findIndex(function (row) { return String(row[0]).trim() === marker; });
    if (start < 0 || !rows[start + 1]) return [];
    var headers = rows[start + 1].map(function (cell) { return String(cell).trim(); });
    var end = nextMarker ? rows.findIndex(function (row, i) {
      return i > start && String(row[0]).trim() === nextMarker;
    }) : rows.length;
    if (end < 0) end = rows.length;
    return rows.slice(start + 2, end).filter(function (row) {
      return row.some(function (cell) { return cell !== ''; }) && String(row[0]).trim() !== '#REF!';
    }).map(function (row) {
      var record = {};
      headers.forEach(function (header, i) { if (header) record[header] = row[i]; });
      return record;
    });
  }

  function parseTournament(wb) {
    if (!wb.Sheets.Export) throw new Error('Nel manager manca il foglio Export.');
    var rows = XLSX.utils.sheet_to_json(wb.Sheets.Export, { header: 1, defval: '', raw: false });
    var standingsStart = rows.findIndex(function (row) { return String(row[0]).trim() === 'standings'; });
    var matchesStart = rows.findIndex(function (row) { return String(row[0]).trim() === 'matches'; });
    var standings = rows.slice(standingsStart + 2, matchesStart).filter(function (row) {
      return row[0] !== '' && row[2] !== '';
    }).map(function (row) {
      return {
        group: row[0], position: row[1], team: row[2], points: row[3], played: row[4],
        won: row[5], won_pen: row[6], lost_pen: row[7], lost: row[8],
        gf: row[9], ga: row[10], gd: row[11]
      };
    });
    var matches = recordsBetween(rows, 'matches', 'knockout');
    var knockout = recordsBetween(rows, 'knockout', 'match_goals');
    var goals = recordsBetween(rows, 'match_goals');
    var matchMap = {};

    matches.forEach(function (match) {
      match.scorers = { home: [], away: [], autogoals: [] };
      matchMap[String(match.match_id)] = match;
    });
    goals.forEach(function (goal) {
      var match = matchMap[String(goal.match_id)];
      if (!match || !goal.goal_type) return;
      var player = String(goal.player || '').replace(/^\s*\d+\s*-\s*/, '').trim();
      if (String(goal.goal_type).toLowerCase() === 'autogol') {
        if (String(goal.goal_for_team) === String(match.home_team)) match.scorers.home.push('Autogol');
        if (String(goal.goal_for_team) === String(match.away_team)) match.scorers.away.push('Autogol');
      } else if (String(goal.goal_for_team) === String(match.home_team) && player) {
        match.scorers.home.push(player);
      } else if (String(goal.goal_for_team) === String(match.away_team) && player) {
        match.scorers.away.push(player);
      }
    });

    var topScorers = [];
    rows.forEach(function (row) {
      var scorerPosition = Number(row[13]);
      var scorerGoals = Number(row[16]);
      if (Number.isFinite(scorerPosition) && scorerPosition > 0 &&
          row[14] !== '' && row[15] !== '' &&
          Number.isFinite(scorerGoals) && scorerGoals > 0) {
        topScorers.push({ position: row[13], player: row[14], team: row[15], goals: row[16] });
      }
    });
    return { standings: standings, matches: matches, knockout: knockout, topScorers: topScorers };
  }

  function rankClass(position) {
    return Number(position) === 1 ? 'first' : Number(position) === 2 ? 'second' : Number(position) === 3 ? 'third' : '';
  }

  function renderStandings(data) {
    ['A', 'B', 'C', 'D'].forEach(function (group) {
      var tbody = document.querySelector('.classifica-table[data-girone="' + group + '"] tbody');
      if (!tbody) return;
      tbody.innerHTML = data.standings.filter(function (team) {
        return String(team.group) === group;
      }).map(function (team) {
        return '<tr><td><span class="team-rank ' + rankClass(team.position) + '">' + esc(team.position) + '</span></td>' +
          '<td>' + esc(team.team) + '</td><td>' + esc(team.played) + '</td><td>' + esc(team.won) + '</td>' +
          '<td>' + esc(team.lost) + '</td><td>' + esc(team.won_pen) + '</td><td>' + esc(team.lost_pen) + '</td>' +
          '<td>' + esc(team.gf) + '</td><td>' + esc(team.ga) + '</td><td>' +
          (Number(team.gd) > 0 ? '+' : '') + esc(team.gd) + '</td><td><span class="pts-badge">' +
          esc(team.points) + '</span></td></tr>';
      }).join('');
    });
    var lastTable = document.querySelector('[data-girone="D"]');
    var standingsSection = lastTable ? lastTable.closest('.risultati-section') : null;
    var notice = standingsSection ? standingsSection.querySelector('.risultati-empty') : null;
    if (notice) notice.style.display = 'none';
  }

  function isLive(match) {
    if (!match.time || String(match.status).toLowerCase() === 'complete') return false;
    var time = String(match.time).match(/^(\d{1,2}):(\d{2})/);
    if (!time) return false;
    var start = new Date(TOURNAMENT_DATE + 'T' + time[1].padStart(2, '0') + ':' + time[2] + ':00+02:00');
    var minutes = String(match.stage).toLowerCase() === 'group' ? 20 : 30;
    var now = new Date();
    return now >= start && now < new Date(start.getTime() + minutes * 60000);
  }

  function scorerList(names) {
    var counts = {};
    names.forEach(function (name) { counts[name] = (counts[name] || 0) + 1; });
    return Object.keys(counts).map(function (name) {
      return esc(name) + (counts[name] > 1 ? ' ×' + counts[name] : '');
    }).join('<br>');
  }

  function matchCard(match, showScorers) {
    var complete = String(match.status).toLowerCase() === 'complete' ||
      (match.home_goals !== '' && match.away_goals !== '');
    var live = isLive(match);
    var score = complete ? esc(match.home_goals) + ' - ' + esc(match.away_goals) : '—';
    if (complete && match.penalty_winner) score += '<br><small>Rigori: ' + esc(match.penalty_winner) + '</small>';
    var group = match.group ? '<span>Girone ' + esc(match.group) + '</span>' : '';
    var meta = '<div class="partita-meta">' + group + '<span>' + esc(match.time) + '</span>' +
      (match.field ? '<span>Campo ' + esc(match.field) + '</span>' : '') +
      (live ? '<span class="live-badge"><span class="live-dot"></span>Live</span>' : '') + '</div>';
    var scorers = '';
    if (showScorers && complete) {
      var home = scorerList(match.scorers.home);
      var away = scorerList(match.scorers.away);
      if (home || away) {
        scorers = '<div class="partita-scorers"><div>' + (home || '—') +
          '</div><div class="away">' + (away || '—') + '</div></div>';
      }
    }
    return '<div class="partita-card' + (live ? ' live' : '') + '">' + meta +
      '<div class="partita-team">' + esc(match.home_team) + '</div><div class="partita-score' +
      (complete ? '' : ' tbd') + '">' + score + '</div><div class="partita-team away">' +
      esc(match.away_team) + '</div>' + scorers + '</div>';
  }

  function renderGroupMatches() {
    if (!tournamentData) return;
    var container = document.getElementById('groupMatches');
    var empty = document.getElementById('groupMatchesEmpty');
    if (!container) return;
    var matches = tournamentData.matches.filter(function (match) {
      return String(match.stage).toLowerCase() === 'group' &&
        (selectedGroup === 'all' || String(match.group) === selectedGroup);
    });
    container.innerHTML = matches.map(function (match) {
      return matchCard(match, selectedGroup !== 'all');
    }).join('');
    if (empty) empty.style.display = matches.length ? 'none' : 'block';
  }

  function renderTopScorers(data) {
    var tbody = document.getElementById('topScorersBody');
    var empty = document.getElementById('topScorersEmpty');
    if (!tbody) return;
    tbody.innerHTML = data.topScorers.map(function (scorer) {
      return '<tr><td><span class="team-rank ' + rankClass(scorer.position) + '">' + esc(scorer.position) +
        '</span></td><td>' + esc(scorer.player) + '</td><td>' + esc(scorer.team) +
        '</td><td><span class="pts-badge">' + esc(scorer.goals) + '</span></td></tr>';
    }).join('');
    if (empty) empty.style.display = data.topScorers.length ? 'none' : 'block';
  }

  function renderKnockout(wb) {
    if (!wb.Sheets.Knockout) return;
    var matches = XLSX.utils.sheet_to_json(wb.Sheets.Knockout, { defval: '', raw: false });
    var targets = { 'Quarter-final': 'quarterFinals', 'Semi-final': 'semiFinals', 'Final': 'tournamentFinal' };
    Object.keys(targets).forEach(function (stage) {
      var target = document.getElementById(targets[stage]);
      if (!target) return;
      target.innerHTML = matches.filter(function (row) { return row.Stage === stage; }).map(function (row) {
        return matchCard({
          stage: row.Stage, group: '', field: row.Field, time: row.Time,
          home_team: row.HomeTeam || 'TBD', away_team: row.AwayTeam || 'TBD',
          home_goals: row.HomeGoals, away_goals: row.AwayGoals, penalty_winner: row.PenaltyWinner,
          status: row.HomeGoals !== '' && row.AwayGoals !== '' ? 'Complete' : 'Pending',
          scorers: { home: [], away: [], autogoals: [] }
        }, false);
      }).join('');
    });
  }

  function setupTabs() {
    var tabs = document.getElementById('risultatiTabs');
    if (!tabs || tabs.dataset.ready) return;
    tabs.dataset.ready = 'true';
    tabs.addEventListener('click', function (event) {
      var button = event.target.closest('.risultati-tab');
      if (!button) return;
      selectedGroup = button.dataset.group;
      tabs.querySelectorAll('.risultati-tab').forEach(function (tab) {
        var active = tab === button;
        tab.classList.toggle('active', active);
        tab.setAttribute('aria-selected', active ? 'true' : 'false');
      });
      renderGroupMatches();
    });
  }

  window.loadTournamentManagerData = function () {
    return loadWorkbook().then(function (wb) {
      tournamentData = tournamentData || parseTournament(wb);
      renderStandings(tournamentData);
      renderTopScorers(tournamentData);
      setupTabs();
      renderGroupMatches();
      renderKnockout(wb);
      clearInterval(liveTimer);
      liveTimer = setInterval(renderGroupMatches, 30000);
    }).catch(function (error) {
      console.error('Errore caricamento 12HBosaro_manager.xlsx:', error);
      var empty = document.getElementById('groupMatchesEmpty');
      if (empty) { empty.textContent = 'Impossibile caricare i dati del torneo.'; empty.style.display = 'block'; }
    });
  };

  window.loadTeamsManagerData = function () {
    return loadWorkbook().then(function (wb) {
      if (!wb.Sheets.Teams || !wb.Sheets.Players) throw new Error('Nel manager mancano Teams o Players.');
      var teams = XLSX.utils.sheet_to_json(wb.Sheets.Teams, { defval: '', raw: false });
      var players = XLSX.utils.sheet_to_json(wb.Sheets.Players, { defval: '', raw: false });
      var container = document.getElementById('squadreContainer');
      if (!container) return;
      container.innerHTML = teams.map(function (team) {
        var roster = players.filter(function (player) { return String(player.TeamID) === String(team.TeamID); });
        var rows = roster.map(function (player) {
          var captain = String(player.Captain).toLowerCase() === 'yes';
          return '<div class="player-row' + (captain ? ' captain' : '') + '"><span class="player-number">' +
            esc(player.ShirtNumber || '—') + '</span><span class="player-name' + (captain ? ' captain' : '') +
            '">' + esc(player.PlayerName || '—') + '</span></div>';
        }).join('');
        return '<div class="squadra-card"><div class="squadra-header"><div class="squadra-name">' +
          esc(team.TeamName) + '</div><div class="squadra-count">Girone ' + esc(team.Group) + ' · ' +
          roster.length + ' giocatori</div></div><div class="squadra-players">' + rows + '</div></div>';
      }).join('');
    }).catch(function (error) { console.error('Errore caricamento squadre dal manager:', error); });
  };

  // Le tab devono restare cliccabili anche se il workbook è momentaneamente irraggiungibile.
  setupTabs();
}());

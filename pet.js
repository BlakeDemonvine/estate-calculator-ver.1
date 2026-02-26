/**
 * 網頁寵物設定檔 (包含所有動物)
 * 預設影格數為：Idle(4), Walk(6), Attack(4), Hurt(2), Death(4)
 * 如果您的某張圖影格數不同，請直接在此修改 `frames` 數值！
 */
const PET_CONFIG = {
  // ================= 猛獸組 (滑鼠靠近會攻擊) =================
  'Dog 1': {
    speed: 1.5, 
    states: {
      idle: { file: 'Idle.png', frames: 4, duration: '0.8s', loop: true },
      walk: { file: 'Walk.png', frames: 6, duration: '0.6s', loop: true },
      interact: { file: 'Attack.png', frames: 4, duration: '0.5s', loop: true },
      hurt: { file: 'Hurt.png', frames: 2, duration: '0.2s', loop: true },
      dead: { file: 'Death.png', frames: 4, duration: '0.8s', loop: false }
    }
  },
  'Dog 2': {
    speed: 1.5, 
    states: {
      idle: { file: 'Idle.png', frames: 4, duration: '0.8s', loop: true },
      walk: { file: 'Walk.png', frames: 6, duration: '0.6s', loop: true },
      interact: { file: 'Attack.png', frames: 4, duration: '0.5s', loop: true },
      hurt: { file: 'Hurt.png', frames: 2, duration: '0.2s', loop: true },
      dead: { file: 'Death.png', frames: 4, duration: '0.8s', loop: false }
    }
  },
  'Cat 1': {
    speed: 1.8, // 貓咪走得比狗稍微輕快一點
    states: {
      idle: { file: 'Idle.png', frames: 4, duration: '0.8s', loop: true },
      walk: { file: 'Walk.png', frames: 6, duration: '0.6s', loop: true },
      interact: { file: 'Attack.png', frames: 4, duration: '0.5s', loop: true },
      hurt: { file: 'Hurt.png', frames: 2, duration: '0.2s', loop: true },
      dead: { file: 'Death.png', frames: 4, duration: '0.8s', loop: false }
    }
  },
  'Cat 2': {
    speed: 1.8, 
    states: {
      idle: { file: 'Idle.png', frames: 4, duration: '0.8s', loop: true },
      walk: { file: 'Walk.png', frames: 6, duration: '0.6s', loop: true },
      interact: { file: 'Attack.png', frames: 4, duration: '0.5s', loop: true },
      hurt: { file: 'Hurt.png', frames: 2, duration: '0.2s', loop: true },
      dead: { file: 'Death.png', frames: 4, duration: '0.8s', loop: false }
    }
  },

  // ================= 小動物組 (滑鼠靠近會驚嚇/快跑) =================
  'Bird 1': {
    speed: 2.5, // 鳥飛得最快
    states: {
      idle: { file: 'Idle.png', frames: 4, duration: '0.8s', loop: true },
      walk: { file: 'Walk.png', frames: 6, duration: '0.6s', loop: true },
      interact: { file: 'Walk.png', frames: 6, duration: '0.3s', loop: true }, // 加速拍翅膀
      hurt: { file: 'Hurt.png', frames: 2, duration: '0.2s', loop: true },
      dead: { file: 'Death.png', frames: 4, duration: '0.8s', loop: false }
    }
  },
  'Bird 2': {
    speed: 2.5, 
    states: {
      idle: { file: 'Idle.png', frames: 4, duration: '0.8s', loop: true },
      walk: { file: 'Walk.png', frames: 6, duration: '0.6s', loop: true },
      interact: { file: 'Walk.png', frames: 6, duration: '0.3s', loop: true }, 
      hurt: { file: 'Hurt.png', frames: 2, duration: '0.2s', loop: true },
      dead: { file: 'Death.png', frames: 4, duration: '0.8s', loop: false }
    }
  },
  'Rat 1': {
    speed: 2.2, // 老鼠竄得也很快
    states: {
      idle: { file: 'Idle.png', frames: 4, duration: '0.8s', loop: true },
      walk: { file: 'Walk.png', frames: 6, duration: '0.6s', loop: true },
      interact: { file: 'Walk.png', frames: 6, duration: '0.2s', loop: true }, // 遇到滑鼠會像無頭蒼蠅一樣快速原地跑
      hurt: { file: 'Hurt.png', frames: 2, duration: '0.2s', loop: true },
      dead: { file: 'Death.png', frames: 4, duration: '0.8s', loop: false }
    }
  },
  'Rat 2': {
    speed: 2.2, 
    states: {
      idle: { file: 'Idle.png', frames: 4, duration: '0.8s', loop: true },
      walk: { file: 'Walk.png', frames: 6, duration: '0.6s', loop: true },
      interact: { file: 'Walk.png', frames: 6, duration: '0.2s', loop: true }, 
      hurt: { file: 'Hurt.png', frames: 2, duration: '0.2s', loop: true },
      dead: { file: 'Death.png', frames: 4, duration: '0.8s', loop: false }
    }
  }
};

/**
 * 創建一隻會動的網頁寵物
 * @param {string} petName - 動物名稱 (需對應 PET_CONFIG)
 */
function createActivePet(petName) {
  const config = PET_CONFIG[petName];
  if (!config) {
    console.error('找不到這隻寵物的設定檔：', petName);
    return;
  }

  // 1. 注入共用的 CSS 動畫
  if (!document.getElementById('web-pet-styles')) {
    const styleNode = document.createElement('style');
    styleNode.id = 'web-pet-styles';
    styleNode.innerHTML = `
      .pet-container {
        position: fixed;
        z-index: 99999;
        width: 48px;    
        height: 48px;   
        pointer-events: auto;
        cursor: pointer;
        image-rendering: pixelated;
        background-repeat: no-repeat;
      }
      @keyframes pet-anim-2 { 100% { background-position: -96px 0px; } }
      @keyframes pet-anim-4 { 100% { background-position: -192px 0px; } }
      @keyframes pet-anim-6 { 100% { background-position: -288px 0px; } }
    `;
    document.head.appendChild(styleNode);
  }

  const petDiv = document.createElement('div');
  petDiv.classList.add('pet-container');
  document.body.appendChild(petDiv);

  let state = '';
  let x = Math.random() * (window.innerWidth - 48);
  let y = Math.random() * (window.innerHeight - 48);
  let targetX = x;
  let targetY = y;
  let direction = 1;
  let interactTimeout = null;

  function setState(newState) {
    if (state === 'dead') return; 
    state = newState;
    const stateConfig = config.states[newState];

    petDiv.style.backgroundImage = `url('pet/${petName}/${stateConfig.file}')`;
    petDiv.style.backgroundSize = `${stateConfig.frames * 48}px 48px`; 

    const animName = `pet-anim-${stateConfig.frames}`;
    const playMode = stateConfig.loop ? 'infinite' : 'forwards';
    
    petDiv.style.animation = 'none';
    petDiv.offsetHeight; // 觸發重繪
    petDiv.style.animation = `${animName} ${stateConfig.duration} steps(${stateConfig.frames}) ${playMode}`;
  }

  setState('idle');

  function updateLoop() {
    if (state !== 'dead' && state !== 'hurt') {
      if (state === 'walk') {
        const dx = targetX - x;
        const dy = targetY - y;
        const dist = Math.sqrt(dx * dx + dy * dy);

        if (dist < config.speed) {
          x = targetX; y = targetY; 
          setState('idle');
        } else {
          x += (dx / dist) * config.speed;
          y += (dy / dist) * config.speed;
          direction = dx > 0 ? 1 : -1; 
        }
      }
      petDiv.style.left = x + 'px';
      petDiv.style.top = y + 'px';
      petDiv.style.transform = `scaleX(${direction})`;
    }
    requestAnimationFrame(updateLoop);
  }
  requestAnimationFrame(updateLoop);

  setInterval(() => {
    if (state === 'dead' || state === 'hurt' || state === 'interact') return;
    if (Math.random() > 0.5) {
      targetX = Math.random() * (window.innerWidth - 48);
      targetY = Math.random() * (window.innerHeight - 48);
      setState('walk');
    } else {
      setState('idle');
    }
  }, 2500);

  document.addEventListener('mousemove', (e) => {
    if (state === 'dead' || state === 'hurt') return;
    const petCenterX = x + 24;
    const petCenterY = y + 24;
    const dist = Math.sqrt((e.clientX - petCenterX) ** 2 + (e.clientY - petCenterY) ** 2);

    if (dist < 80) {
      if (state !== 'interact') setState('interact');
      direction = e.clientX - petCenterX > 0 ? 1 : -1;
      petDiv.style.transform = `scaleX(${direction})`;

      clearTimeout(interactTimeout);
      interactTimeout = setTimeout(() => {
        if (state === 'interact') setState('idle');
      }, 800);
    }
  });

  petDiv.addEventListener('click', (e) => {
    e.stopPropagation();
    if (state === 'dead' || state === 'hurt') return;
    
    clearTimeout(interactTimeout);
    petDiv.style.cursor = 'default';

    setState('hurt');
    setTimeout(() => {
      setState('dead');
    }, 400); 
  });
}

// ================= 在網頁載入時，把所有動物都召喚出來！ =================
/*
window.addEventListener('load', () => {
  createActivePet('Dog 1');
  createActivePet('Dog 2');
  createActivePet('Cat 1');
  createActivePet('Cat 2');
  createActivePet('Bird 1');
  createActivePet('Bird 2');
  createActivePet('Rat 1');
  createActivePet('Rat 2');
});
*/
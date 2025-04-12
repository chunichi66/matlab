% サンプル信号（例としてsin波＋ノイズ）
fs = 1000; % サンプリング周波数 (Hz)
t = 0:1/fs:2; % 時間ベクトル（2秒間）
x = sin(2*pi*100*t) + 0.5*sin(2*pi*200*t) + 0.2*randn(size(t)); % 複合信号

% スペクトログラム描画
figure;
pspectrum(x, fs, 'spectrogram', ...
    'OverlapPercent', 80, ...
    'FrequencyLimits',[1 300], ...
    'MinThreshold', -30);
colorbar;
title('スペクトログラム（高パワー＝暗）');
xlabel('時間 (秒)');
ylabel('周波数 (Hz)');

% カラーマップ：白→赤→黒（値が大きいほど暗い）
n = 256;
half_n = round(n/2);

% 赤→黒（高パワー帯）
red_to_black = [ ...
    linspace(1,0,half_n)', ...
    zeros(half_n,1), ...
    zeros(half_n,1)
];

% 白→赤（低パワー帯）
white_to_red = [ ...
    ones(half_n,1), ...
    linspace(1,0,half_n)', ...
    linspace(1,0,half_n)'
];

custom_cmap = [white_to_red; red_to_black];

% カラーマップ適用
colormap(custom_cmap);


% パワースペクトル計算
[ps, f] = pspectrum(x, fs);

% dBスケールに変換
ps_dB = 10*log10(ps);

% プロット
figure;
plot(f, ps_dB, 'LineWidth', 1.5);
xlabel('周波数 (Hz)');
ylabel('パワー (dB)');
title('パワースペクトル');
grid on;


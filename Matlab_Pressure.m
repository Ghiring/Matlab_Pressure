%%Évaluation des pressions issues de la nappe de pression de l'entreprise LimGroup
%Importation du fichier et partitionnement de la selle
clear;
clc;
clf; 

File = readmatrix("FichierLimGroup.xlsx");                  
Selle = File(3:end,19:242);                                

    % Partitionnement de la selle

% Cellule 17 = colonne 1
% Cellule 240 = colonne 224
Avant = Selle(:,1:80);                                     
Milieu = Selle(:,81:144);                                 
Arriere = Selle(:,145:224);                                

i=1;
j=8;
Droite=[];
while j<=216
    SelleA = Selle(:,i:j) ;
    j=j+16;
    i=i+16;
    Droite = [Droite,SelleA];
end

i=9;
j=16;
Gauche=[];
while j<=224
    SelleA = Selle(:,i:j) ;
    j=j+16;
    i=i+16;
    Gauche = [Gauche,SelleA];
end

%%Pmoyennes et Pmax à l'image 2200 pusi à toutes les images

Selle2200 = Selle(2200,:);                              
Droite2200 = Droite(2200,:);                              
Gauche2200 = Gauche(2200,:);                                
Milieu2200 = Milieu(2200,:);                                
Avant2200 = Avant(2200,:);                                  
Arriere2200 = Arriere(2200,:);                              

Pmoy2200_selle = mean(Selle2200);
Pmoy2200_arriere = mean(Arriere2200);
Pmoy2200_avant = mean(Avant2200);
Pmoy2200_milieu = mean(Milieu2200);                         
Pmoy2200_gauche = mean(Gauche2200);  
Pmoy2200_droite = mean(Droite2200);

Pmax2200_selle = max(Selle2200);
Pmax2200_avant = max(Avant2200);    
Pmax2200_arriere = max(Arriere2200);                        
Pmax2200_milieu = max(Milieu2200);
Pmax2200_gauche = max(Gauche2200);
Pmax2200_droite = max(Droite2200);

Pmoy_selle = mean(mean(Selle));
Pmoy_arriere = mean(mean(Arriere));
Pmoy_avant = mean(mean(Avant));
Pmoy_milieu = mean(mean(Milieu));                   
Pmoy_gauche = mean(mean(Gauche));  
Pmoy_droite = mean(mean(Droite));

Pmax_selle = max(max(Selle));
Pmax_avant = max(max(Avant));
Pmax_arriere = max(max(Arriere));                   
Pmax_milieu = max(max(Milieu));
Pmax_gauche = max(max(Gauche));
Pmax_droite = max(max(Droite));

%%Calcul des COP à toutes les images (y compris 2200)
    
    % Partie Droite

COPxglobD=0;
COPyglobD=0;
MatCOPxglobD = [];
MatCOPyglobD = [];
y = 1;
z=1;
a=1;
v=1;
x = reshape(Droite(y,:),14,[]);

while y <= (height(Droite)-1)            

    while z<=14
        COPxD = ((1*x(z,1))+(2*x(z,2))+(3*x(z,3))+(4*x(z,4))+(5*x(z,5))+(6*x(z,6))+(7*x(z,7))+(8*x(z,8)));
        COPxD(isnan(COPxD))=0;
        COPxD = COPxD/sum(x(z,:));
        COPxglobD = COPxglobD + COPxD; 
        MatCOPxglobD = [MatCOPxglobD , COPxD];
        z=z+1;    
    end 
    
    while a<=8
        COPyD = (1*x(1,a))+(2*x(2,a))+(3*x(3,a))+(4*x(4,a))+(5*x(5,a))+(6*x(6,a))+(7*x(7,a))+(8*x(8,a)) ...
            +(9*x(9,a))+(10*x(10,a))+(11*x(11,a))+(12*x(12,a))+(13*x(13,a))+(14*x(14,a));
        COPyD(isnan(COPyD))=0;
        COPyD = COPyD/sum(x(:,a));
        COPyglobD = COPyglobD + COPyD;
        MatCOPyglobD =  [MatCOPyglobD , COPyD];
        a=a+1;

    end

    if y==2200
            COPx2200D = COPxD
            COPy2200D = COPyD
    end

    y=y+1;
end 
MatCOPxglobD = MatCOPxglobD';
COPxSelleD = mean(MatCOPxglobD)
COPySelleD = mean(MatCOPyglobD)
   
   % Partie Gauche

COPxglobG=0;
COPyglobG=0;
MatCOPxglobG = [];
MatCOPyglobG = [];
y = 1;
z=1;
a=1;
v=1;
x = reshape(Gauche(y,:),14,[]);

while y <= (height(Gauche)-1)            

    while z<=14
        COPxG = ((1*x(z,1))+(2*x(z,2))+(3*x(z,3))+(4*x(z,4))+(5*x(z,5))+(6*x(z,6))+(7*x(z,7))+(8*x(z,8)));
        COPxG(isnan(COPxG))=0;
        COPxG = COPxG/sum(x(z,:));
        COPxglobG = COPxglobG + COPxG; 
        MatCOPxglobG = [MatCOPxglobG , COPxG];
        z=z+1;    
    end 
    
    while a<=8
        COPyG = (1*x(1,a))+(2*x(2,a))+(3*x(3,a))+(4*x(4,a))+(5*x(5,a))+(6*x(6,a))+(7*x(7,a))+(8*x(8,a)) ...
            +(9*x(9,a))+(10*x(10,a))+(11*x(11,a))+(12*x(12,a))+(13*x(13,a))+(14*x(14,a));
        COPyG(isnan(COPyG))=0;
        COPyG = COPyG/sum(x(:,a));
        COPyglobG = COPyglobG + COPyG;
        MatCOPyglobG =  [MatCOPyglobG , COPyG];
        a=a+1;

    end

    if y==2200
            COPx2200G = COPxG;
            COPy2200G = COPyG;
    end

    y=y+1;
end 
MatCOPyglobG(isnan(MatCOPyglobG))=0;
MatCOPxglobG = MatCOPxglobG';
COPxSelleG = mean(MatCOPxglobG);
COPySelleG = mean(MatCOPyglobG);
    
    % Partie Avant

COPxglobAv=0;
COPyglobAv=0;
MatCOPxglobAv = [];
MatCOPyglobAv = [];
y = 1;
z=1;
a=1;
v=1;
x = reshape(Avant(y,:),5,[]);

while y <= (height(Avant)-1)            

    while z<=5
        COPxAv = ((1*x(z,1))+(2*x(z,2))+(3*x(z,3))+(4*x(z,4))+(5*x(z,5))+(6*x(z,6))+(7*x(z,7))+(8*x(z,8))...
            +(9*x(z,9))+(10*x(z,10))+(11*x(z,11))+(12*x(z,12))+(13*x(z,13))+(14*x(z,14))+(15*x(z,15))+(16*x(z,16)));
        COPxAv(isnan(COPxAv))=0;
        COPxAv = COPxAv/sum(x(z,:));
        COPxglobAv = COPxglobAv + COPxAv; 
        MatCOPxglobAv = [MatCOPxglobAv , COPxAv];
        z=z+1;    
    end 
    
    while a<=16
        COPyAv = (1*x(1,a))+(2*x(2,a))+(3*x(3,a))+(4*x(4,a))+(5*x(5,a));
        COPyAv(isnan(COPyAv))=0;
        COPyAv = COPyAv/sum(x(:,a));
        COPyglobAv = COPyglobAv + COPyAv;
        MatCOPyglobAv =  [MatCOPyglobAv , COPyAv];
        a=a+1;

    end

    if y==2200
            COPx2200Av = COPxAv;
            COPy2200Av = COPyAv;
    end

    y=y+1;
end 
MatCOPyglobAv(isnan(MatCOPyglobAv))=0;
MatCOPxglobAv = MatCOPxglobAv';
COPxSelleAv = mean(MatCOPxglobAv);
COPySelleAv = mean(MatCOPyglobAv);
   
   % Partie Arriere

COPxglobAr=0;
COPyglobAr=0;
MatCOPxglobAr = [];
MatCOPyglobAr = [];
y = 1;
z=1;
a=1;
v=1;
x = reshape(Arriere(y,:),5,[]);

while y <= (height(Arriere)-1)            

    while z<=5
        COPxAr = ((1*x(z,1))+(2*x(z,2))+(3*x(z,3))+(4*x(z,4))+(5*x(z,5))+(6*x(z,6))+(7*x(z,7))+(8*x(z,8))...
            +(9*x(z,9))+(10*x(z,10))+(11*x(z,11))+(12*x(z,12))+(13*x(z,13))+(14*x(z,14))+(15*x(z,15))+(16*x(z,16)));
        COPxAr(isnan(COPxAr))=0;
        COPxAr = COPxAr/sum(x(z,:));
        COPxglobAr = COPxglobAr + COPxAr; 
        MatCOPxglobAr = [MatCOPxglobAr , COPxAr];
        z=z+1;    
    end 
    
    while a<=16
        COPyAr = (1*x(1,a))+(2*x(2,a))+(3*x(3,a))+(4*x(4,a))+(5*x(5,a));
        COPyAr(isnan(COPyAr))=0;
        COPyAr = COPyAr/sum(x(:,a));
        COPyglobAr = COPyglobAr + COPyAr;
        MatCOPyglobAr =  [MatCOPyglobAr , COPyAr];
        a=a+1;

    end

    if y==2200
            COPx2200Ar = COPxAr;
            COPy2200Ar = COPyAr;
    end

    y=y+1;
end 
MatCOPyglobAr(isnan(MatCOPyglobAr))=0;
MatCOPxglobAr = MatCOPxglobAr';
COPxSelleAr = mean(MatCOPxglobAr);
COPySelleAr = mean(MatCOPyglobAr);
  
  % Partie Centrale

COPxglobM=0;
COPyglobM=0;
MatCOPxglobM = [];
MatCOPyglobM = [];
y = 1;
z=1;
a=1;
v=1;
x = reshape(Milieu(y,:),4,[]);

while y <= (height(Milieu)-1)            

    while z<=4
        COPxM = ((1*x(z,1))+(2*x(z,2))+(3*x(z,3))+(4*x(z,4))+(5*x(z,5))+(6*x(z,6))+(7*x(z,7))+(8*x(z,8))...
            +(9*x(z,9))+(10*x(z,10))+(11*x(z,11))+(12*x(z,12))+(13*x(z,13))+(14*x(z,14))+(15*x(z,15))+(16*x(z,16)));
        COPxM(isnan(COPxM))=0;
        COPxM = COPxM/sum(x(z,:));
        COPxglobM = COPxglobM + COPxM; 
        MatCOPxglobM = [MatCOPxglobM , COPxM];
        z=z+1;    
    end 
    
    while a<=16
        COPyM = (1*x(1,a))+(2*x(2,a))+(3*x(3,a))+(4*x(4,a));
        COPyM(isnan(COPyM))=0;
        COPyM = COPyM/sum(x(:,a));
        COPyglobM = COPyglobM + COPyM;
        MatCOPyglobM =  [MatCOPyglobM , COPyM];
        a=a+1;

    end

    if y==2200
        COPx2200M = COPxM;
        COPy2200M = COPyM;
    end

    y=y+1;
end 
MatCOPxglobM(isnan(MatCOPxglobM))=0;
MatCOPyglobM(isnan(MatCOPyglobM))=0;
MatCOPxglobM = MatCOPxglobM';
COPxSelleM = mean(MatCOPxglobM);
COPySelleM = mean(MatCOPyglobM);
    
    % Selle entière

COPxglob=0;
COPyglob=0;
MatCOPxglob = [];
MatCOPyglob = [];
y = 1;
z=1;
a=1;
v=1;
x = reshape(Selle(y,:),14,[]);

while y <= (height(Selle)-1)            

    while z<=14
        COPx = ((1*x(z,1))+(2*x(z,2))+(3*x(z,3))+(4*x(z,4))+(5*x(z,5))+(6*x(z,6))+(7*x(z,7))+(8*x(z,8)) ...
            +(9*x(z,9))+(10*x(z,10))+(11*x(z,11))+(12*x(z,12))+(13*x(z,13))+(14*x(z,14))+(15*x(z,15))+(16*x(z,16)));
        COPx(isnan(COPx))=0;
        COPx = COPx/sum(x(z,:));
        COPxglob = COPxglob + COPx; 
        MatCOPxglob = [MatCOPxglob , COPx];
        z=z+1;    
    end 
    
    while a<=16
        COPy = (1*x(1,a))+(2*x(2,a))+(3*x(3,a))+(4*x(4,a))+(5*x(5,a))+(6*x(6,a))+(7*x(7,a))+(8*x(8,a)) ...
            +(9*x(9,a))+(10*x(10,a))+(11*x(11,a))+(12*x(12,a))+(13*x(13,a))+(14*x(14,a));
        COPy(isnan(COPy))=0;
        COPy = COPy/sum(x(:,a));
        COPyglob = COPyglob + COPy;
        MatCOPyglob =  [MatCOPyglob , COPy];
        a=a+1;

    end

    if y==2200
            COPx2200 = COPx;
            COPy2200 = COPy;
    end

    y=y+1;
end 
MatCOPxglob = MatCOPxglob';
COPxSelle = mean(MatCOPxglob);
COPySelle = mean(MatCOPyglob);

%%Graph3D des pressions à l'image 2200

clf;   
SE = reshape(Selle(2200,:),14,[]);
Graph3D = surf(SE);                                                         
view([201.6 52.9])
colorbar north                                                   
colormap hot                                                                
title("Graph présentant la pression sur la selle à l'image 2200")

%%Calcul des aires à l'image 2200

% Selle entière

Moy5E = 0.05*Pmoy2200_selle;                                 
VE = find(Selle2200>Moy5E);                                   
AireE = length(VE)           

% Avant

Moy5Av = 0.05*Pmoy2200_avant;                                 
VAv = find(Avant2200>Moy5Av);                                   
AireAv = length(VAv)        

% Arriere

Moy5Ar = 0.05*Pmoy2200_arriere;                           
VAr = find(Arriere2200>Moy5Ar);
AireAr = length(VAr)

% Milieu

Moy5M = 0.05*Pmoy2200_milieu;
VM = find(Milieu2200>Moy5M);
AireM = length(VM)

% Droite

Moy5M = 0.05*Pmoy2200_droite;
VM = find(Droite2200>Moy5M);
AireD = length(VM)

% Gauche

Moy5G = 0.05*Pmoy2200_gauche;
VG = find(Gauche2200>Moy5G);
AireG = length(VG)

%%Corrélations croisées

CorrCroisX = xcorr(MatCOPxglobD,MatCOPxglobG);
plot(CorrCroisX)
CorrCroisY = xcorr(MatCOPyglobD,MatCOPyglobG);
plot(CorrCroisY)

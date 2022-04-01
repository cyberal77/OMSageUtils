object frmSageCredentials: TfrmSageCredentials
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'frmSageCredentials'
  ClientHeight = 110
  ClientWidth = 378
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  DesignSize = (
    378
    110)
  TextHeight = 15
  object lblUsername: TLabel
    Left = 8
    Top = 16
    Width = 59
    Height = 15
    Caption = 'Utilisateur :'
  end
  object lblPassword: TLabel
    Left = 8
    Top = 45
    Width = 76
    Height = 15
    Caption = 'Mot de passe :'
  end
  object btnOk: TButton
    Left = 214
    Top = 77
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'Ok'
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = btnOkClick
  end
  object btnCancel: TButton
    Left = 295
    Top = 77
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Cancel = True
    Caption = 'Annuler'
    ModalResult = 2
    TabOrder = 1
  end
  object edtUsername: TEdit
    Left = 89
    Top = 13
    Width = 281
    Height = 23
    TabOrder = 2
  end
  object edtPassword: TEdit
    Left = 89
    Top = 42
    Width = 281
    Height = 23
    PasswordChar = '*'
    TabOrder = 3
  end
end
